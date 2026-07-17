/**
 * ПАРАМЕТРИЧЕСКИЙ БЮДЖЕТ ТРАНСПОРТНЫХ РАСХОДОВ
 * Скрипт сам генерирует тестовые исходные данные (Vehicles, Customers, Seasonality, Params)
 * и строит на их основе: Бюджет, Сценарии (чувствительность к изменению выручки), Дашборд.
 *
 * Три независимых рычага роста выручки (лист Params), каждый со своим влиянием на расходы:
 *   - revenue_scenario_pct — рост объёма у уже существующих клиентов (расходы растут пропорционально)
 *   - customer_count_growth_pct — рост/отток числа клиентов (реально добавляет/убирает рейсы и км)
 *   - price_growth_pct — рост цены за кг (выручка растёт, расходы — нет: чистая маржа)
 *
 * Плюс два рычага на стороне затрат:
 *   - additional_vehicles_light/medium/heavy — докупка машин; лист "CAPEX vs Аутсорс" считает,
 *     окупается ли покупка экономией на аутсорсе
 *   - outsource_tier2/3_threshold_km_month + outsource_tier2/3_rate_rub_per_km — скидка 3PL за объём:
 *     чем больше суммарный аутсорс-пробег компании в месяц, тем ниже ставка за км
 *
 * Как подключить:
 *  1. В Google Sheets: Расширения -> Apps Script.
 *  2. Удалите содержимое Code.gs и вставьте этот файл целиком.
 *  3. Сохраните проект (значок дискеты) и обновите страницу таблицы.
 *  4. В таблице появится меню «Транспортный бюджет»:
 *       0. Сгенерировать тестовые данные — создаёт Vehicles/Customers/Seasonality/Params
 *       1. Пересчитать бюджет — помесячный бюджет на листе «Бюджет»
 *       2. Сценарный анализ выручки — чувствительность затрат к изменению выручки («Сценарии»)
 *       3. Разбивка по видам транспорта, регионам, клиентам, дистанции, сегментам и CAPEX — листы
 *          «По видам транспорта», «По регионам», «Топ клиентов», «По дистанции», «По сегментам
 *          клиентов», «CAPEX vs Аутсорс»
 *       4. Построить дашборд — графики на листе «Дашборд»
 *     Либо один клик: «Собрать всё с нуля».
 */

var SHEET_VEHICLES = 'Vehicles';
var SHEET_CUSTOMERS = 'Customers';
var SHEET_SEASONALITY = 'Seasonality';
var SHEET_PARAMS = 'Params';
var SHEET_BUDGET = 'Бюджет';
var SHEET_SCENARIOS = 'Сценарии';
var SHEET_TRANSPORT_TYPES = 'По видам транспорта';
var SHEET_REGIONS = 'По регионам';
var SHEET_TOP_CUSTOMERS = 'Топ клиентов';
var SHEET_DISTANCE_BANDS = 'По дистанции';
var SHEET_SEGMENTS = 'По сегментам клиентов';
var SHEET_CAPEX = 'CAPEX vs Аутсорс';
var SHEET_DASHBOARD = 'Дашборд';

var TOP_CUSTOMERS_N = 20;

var CATEGORY_LABELS = ['Постоянные расходы', 'Топливо', 'ТО и ремонт', 'Платные дороги',
  'Доплата водителям (сдельная)', 'Суточные / командировочные', 'Аутсорс (3PL)'];

var TIER_LABELS = { Light: 'Лёгкий (Газели, до 3т)', Medium: 'Средний (5-10т)', Heavy: 'Тяжёлый (фуры, 15-20т)' };
var TIER_ORDER = ['Light', 'Medium', 'Heavy'];

var DISTANCE_BANDS = [
  { label: '0–100 км', min: 0, max: 100 },
  { label: '100–300 км', min: 100, max: 300 },
  { label: '300–600 км', min: 300, max: 600 },
  { label: '600–1000 км', min: 600, max: 1000 },
  { label: '1000–2000 км', min: 1000, max: Infinity }
];
function bandForDistance_(d) {
  for (var i = 0; i < DISTANCE_BANDS.length; i++) {
    if (d >= DISTANCE_BANDS[i].min && d < DISTANCE_BANDS[i].max) return DISTANCE_BANDS[i].label;
  }
  return DISTANCE_BANDS[DISTANCE_BANDS.length - 1].label;
}
var SEGMENT_ORDER = ['Розница', 'Опт', 'Крупный дистрибьютор'];

var REGIONS = ['Москва', 'Московская обл.', 'Санкт-Петербург', 'Ленинградская обл.',
  'Казань', 'Нижний Новгород', 'Екатеринбург', 'Самара', 'Новосибирск',
  'Ростов-на-Дону', 'Краснодар', 'Воронеж', 'Уфа', 'Пермь', 'Волгоград',
  'Челябинск', 'Красноярск', 'Владивосток', 'Иркутск', 'Омск'];

// ---------------------------------------------------------------------------
// МЕНЮ
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Транспортный бюджет')
    .addItem('Собрать всё с нуля (данные + бюджет + сценарии + дашборд)', 'runAllFromScratch')
    .addSeparator()
    .addItem('0. Сгенерировать тестовые данные', 'generateSampleData')
    .addItem('1. Пересчитать бюджет', 'buildBudget')
    .addItem('2. Сценарный анализ выручки', 'runScenarioAnalysis')
    .addItem('3. Разбивка по видам транспорта, регионам, клиентам, дистанции и сегментам', 'buildTransportBreakdown')
    .addItem('4. Построить дашборд', 'buildDashboard')
    .addToUi();
}

function runAllFromScratch() {
  generateSampleData();
  runAll();
}

function runAll() {
  buildBudget();
  runScenarioAnalysis();
  buildTransportBreakdown();
  buildDashboard();
  SpreadsheetApp.getUi().alert('Готово: бюджет, сценарии, разбивка по видам/регионам и дашборд обновлены.');
}

// ---------------------------------------------------------------------------
// ГЕНЕРАЦИЯ ТЕСТОВЫХ ДАННЫХ
// ---------------------------------------------------------------------------
function generateSampleData() {
  populateVehicles_();
  populateSeasonality_();
  populateParams_();
  populateCustomers_(1000);
  SpreadsheetApp.getUi().alert('Тестовые данные созданы: Vehicles (13), Seasonality (12 мес.), Params, Customers (1000).');
}

function clearAndGetSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) { sh.clear(); } else { sh = ss.insertSheet(name); }
  return sh;
}

function populateVehicles_() {
  var sh = clearAndGetSheet_(SHEET_VEHICLES);
  var header = ['vehicle_id', 'model', 'tier', 'capacity_t', 'fuel_norm_l_100km',
    'depreciation_month_rub', 'insurance_month_rub', 'driver_fixed_salary_month_rub',
    'variable_driver_rate_rub_km', 'maintenance_rate_rub_km', 'max_daily_km'];
  var rows = [
    [1, 'ГАЗель NEXT', 'Light', 1.5, 12.5, 16000, 4500, 40000, 3.0, 2.3, 250],
    [2, 'ГАЗель NEXT', 'Light', 1.5, 12.5, 16000, 4500, 40000, 3.0, 2.3, 250],
    [3, 'Ford Transit', 'Light', 2.0, 13.5, 18000, 4800, 41000, 3.2, 2.5, 250],
    [4, 'Ford Transit', 'Light', 2.0, 13.5, 18000, 4800, 41000, 3.2, 2.5, 250],
    [5, 'Hyundai HD78', 'Light', 2.5, 15.0, 20000, 5200, 42000, 3.4, 2.7, 260],
    [6, 'Isuzu NQR', 'Medium', 5.0, 18.0, 34000, 8500, 48000, 4.0, 3.8, 300],
    [7, 'Isuzu NQR', 'Medium', 5.0, 18.0, 34000, 8500, 48000, 4.0, 3.8, 300],
    [8, 'Isuzu Forward', 'Medium', 7.0, 20.0, 38000, 9000, 50000, 4.3, 4.1, 300],
    [9, 'MAN TGM', 'Medium', 8.0, 22.0, 41000, 9500, 51000, 4.5, 4.3, 320],
    [10, 'MAN TGM', 'Medium', 10.0, 24.0, 45000, 10000, 53000, 4.8, 4.6, 320],
    [11, 'Volvo FH (тягач)', 'Heavy', 15.0, 28.0, 72000, 14000, 62000, 5.6, 6.0, 550],
    [12, 'Scania R (тягач)', 'Heavy', 18.0, 31.0, 82000, 15500, 65000, 6.0, 6.5, 550],
    [13, 'Kamaz 65117', 'Heavy', 20.0, 33.0, 90000, 17000, 68000, 6.4, 7.0, 500]
  ];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
}

function populateSeasonality_() {
  var sh = clearAndGetSheet_(SHEET_SEASONALITY);
  var header = ['month_num', 'month_name', 'demand_index', 'fuel_price_index'];
  var rows = [
    [1, 'Январь', 0.85, 1.03], [2, 'Февраль', 0.90, 1.02], [3, 'Март', 0.95, 1.00],
    [4, 'Апрель', 1.00, 0.99], [5, 'Май', 1.00, 1.00], [6, 'Июнь', 0.95, 1.01],
    [7, 'Июль', 0.90, 1.02], [8, 'Август', 0.95, 1.02], [9, 'Сентябрь', 1.05, 1.00],
    [10, 'Октябрь', 1.10, 1.00], [11, 'Ноябрь', 1.25, 1.01], [12, 'Декабрь', 1.40, 1.04]
  ];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
}

function populateParams_() {
  var sh = clearAndGetSheet_(SHEET_PARAMS);
  var header = ['param', 'value', 'description'];
  var rows = [
    ['fuel_price_rub_per_l', 62, 'Базовая цена топлива, руб/л'],
    ['perdiem_rate_rub_per_night', 2500, 'Суточные/командировочные за одну ночёвку в рейсе'],
    ['avg_daily_driving_km', 540, 'Средний дневной пробег с учётом норм труда и отдыха (60 км/ч * 9 ч)'],
    ['toll_free_km', 50, 'Первые N км от склада без платных дорог'],
    ['toll_rate_rub_per_km', 1.2, 'Тариф платных дорог, руб/км сверх toll_free_km'],
    ['outsource_rate_rub_per_km', 45, 'Базовая ставка стороннего перевозчика (3PL), руб/км, при превышении мощности парка (1-й тариф)'],
    ['outsource_tier2_threshold_km_month', 200000, 'Порог суммарного аутсорс-пробега компании, км/мес, с которого действует скидочный 2-й тариф'],
    ['outsource_tier2_rate_rub_per_km', 42, 'Ставка 3PL при аутсорс-пробеге выше outsource_tier2_threshold_km_month, руб/км'],
    ['outsource_tier3_threshold_km_month', 350000, 'Порог суммарного аутсорс-пробега компании, км/мес, с которого действует ещё более выгодный 3-й тариф'],
    ['outsource_tier3_rate_rub_per_km', 38, 'Ставка 3PL при аутсорс-пробеге выше outsource_tier3_threshold_km_month, руб/км'],
    ['work_days_per_month', 22, 'Рабочих дней в месяце'],
    ['load_utilization', 0.85, 'Коэффициент использования грузоподъёмности'],
    ['overhead_monthly_fixed_rub', 350000, 'Постоянные накладные расходы логистики (диспетчеризация, GPS-мониторинг, ФОТ логистов)'],
    ['revenue_scenario_pct', 0, 'Рост (+) / падение (-) объёма заказов у УЖЕ существующих клиентов, % (сдельная составляющая роста)'],
    ['customer_count_growth_pct', 0, 'Рост (+) числа клиентов или отток (-), % — добавляет/убирает клиентов с теми же профилями сегментов, что и текущая база'],
    ['price_growth_pct', 0, 'Рост цены за кг груза (пересмотр прайса), % — увеличивает выручку БЕЗ роста пробега и расходов'],
    ['additional_vehicles_light', 0, 'Сколько докупить лёгких машин сверх текущих 5 — сравнение с аутсорсом см. на листе "CAPEX vs Аутсорс"'],
    ['additional_vehicles_medium', 0, 'Сколько докупить средних машин сверх текущих 5'],
    ['additional_vehicles_heavy', 0, 'Сколько докупить тяжёлых машин (фур) сверх текущих 3']
  ];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
}

function populateCustomers_(nCustomers) {
  var sh = clearAndGetSheet_(SHEET_CUSTOMERS);
  var header = ['customer_id', 'customer_name', 'region', 'distance_km', 'segment',
    'avg_order_weight_kg', 'orders_per_month', 'price_per_kg_rub', 'base_monthly_revenue_rub', 'tier'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);

  var rows = [];
  for (var i = 1; i <= nCustomers; i++) {
    var c = generateOneCustomer_(i);
    rows.push([c.customer_id, c.customer_name, c.region, c.distance_km, c.segment,
      c.avg_order_weight_kg, c.orders_per_month, c.price_per_kg_rub, c.base_monthly_revenue_rub, c.tier]);
  }
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

// Детерминированный ГПСЧ (mulberry32): одинаковый seed всегда даёт одну и ту же
// последовательность. Нужен, чтобы состав "синтетических" клиентов при заданном
// customer_count_growth_pct был воспроизводимым — одинаковым во всех листах прогона
// и стабильным между пересчётами (см. buildCustomerList_).
function makeSeededRandom_(seed) {
  var s = seed >>> 0;
  return function () {
    s = (s + 0x6D2B79F5) | 0;
    var t = Math.imul(s ^ (s >>> 15), 1 | s);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// Генерирует одного клиента со случайным профилем — используется и при первичной генерации
// тестовых данных (populateCustomers_), и при моделировании роста числа клиентов
// (buildCustomerList_), чтобы новые "синтетические" клиенты не отличались от исходных.
// rand — источник случайности (по умолчанию Math.random); buildCustomerList_ передаёт
// сюда детерминированный ГПСЧ, чтобы когорта роста была воспроизводимой.
function generateOneCustomer_(id, rand) {
  rand = rand || Math.random;
  function uniform(lo, hi) { return lo + rand() * (hi - lo); }

  function pickSegment() {
    var r = rand();
    if (r < 0.70) return 'Розница';
    if (r < 0.95) return 'Опт';
    return 'Крупный дистрибьютор';
  }

  // Розница — локальная/региональная доставка (мелкие партии не возят за 2000 км).
  // Опт — региональная/межрегиональная. Крупный дистрибьютор — дальнобойные маршруты
  // (крупные партии оправдывают дальние перевозки).
  function pickDistance(segment) {
    var r = rand();
    if (segment === 'Розница') {
      if (r < 0.75) return uniform(5, 100);
      if (r < 0.95) return uniform(100, 300);
      return uniform(300, 600);
    } else if (segment === 'Опт') {
      if (r < 0.40) return uniform(20, 150);
      if (r < 0.80) return uniform(150, 500);
      return uniform(500, 900);
    } else {
      if (r < 0.30) return uniform(100, 400);
      if (r < 0.65) return uniform(400, 1000);
      return uniform(1000, 2000);
    }
  }

  function segmentProfile(segment) {
    if (segment === 'Розница') return { w: [30, 300], o: [2, 4], p: [260, 420] };
    if (segment === 'Опт') return { w: [300, 3000], o: [1, 3], p: [180, 280] };
    return { w: [3000, 12000], o: [1, 2], p: [120, 190] };
  }

  function tierForWeight(w) {
    if (w < 300) return 'Light';
    if (w < 3000) return 'Medium';
    return 'Heavy';
  }

  var segment = pickSegment();
  var prof = segmentProfile(segment);
  var weight = Math.round(uniform(prof.w[0], prof.w[1]) * 10) / 10;
  var orders = Math.floor(uniform(prof.o[0], prof.o[1] + 1));
  var price = Math.round(uniform(prof.p[0], prof.p[1]) * 10) / 10;
  var distance = Math.round(pickDistance(segment) * 10) / 10;
  var revenue = Math.round(weight * orders * price);
  var region = REGIONS[Math.floor(rand() * REGIONS.length)];
  var tier = tierForWeight(weight);
  var name = 'Клиент ' + ('0000' + id).slice(-4);

  return {
    customer_id: id, customer_name: name, region: region, distance_km: distance, segment: segment,
    avg_order_weight_kg: weight, orders_per_month: orders, price_per_kg_rub: price,
    base_monthly_revenue_rub: revenue, tier: tier
  };
}

// Строит список клиентов для расчёта бюджета с учётом параметра customer_count_growth_pct:
// при росте — добавляет синтетических клиентов с теми же профилями сегментов; при оттоке —
// убирает часть существующих. Случайность детерминированно засеяна значением growthPct, поэтому
// при одном и том же customer_count_growth_pct когорта всегда одинакова — состав клиентов
// совпадает во всех листах прогона (бюджет / сценарии / разбивки) и не "плывёт" от того,
// сколько раз функция вызвана. Меняется только при изменении самого параметра роста.
function buildCustomerList_(params) {
  var customers = getCustomers_();
  var growthPct = Number(params.customer_count_growth_pct) || 0;
  if (growthPct === 0) return customers;
  // Seed зависит только от growthPct (масштабируем, чтобы дробные проценты тоже различались).
  var rand = makeSeededRandom_(0x9E3779B9 ^ Math.round(growthPct * 1000));
  if (growthPct > 0) {
    var toAdd = Math.round(customers.length * growthPct / 100);
    var maxId = customers.reduce(function (m, c) { return Math.max(m, c.customer_id); }, 0);
    for (var i = 1; i <= toAdd; i++) {
      customers.push(generateOneCustomer_(maxId + i, rand));
    }
  } else {
    var toRemove = Math.min(customers.length, Math.round(customers.length * (-growthPct) / 100));
    for (var j = 0; j < toRemove; j++) {
      var idx = Math.floor(rand() * customers.length);
      customers.splice(idx, 1);
    }
  }
  return customers;
}

// ---------------------------------------------------------------------------
// ЧТЕНИЕ ИСХОДНЫХ ДАННЫХ
// ---------------------------------------------------------------------------
function getParams_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PARAMS);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var p = {};
  data.forEach(function (r) {
    if (r[0]) p[r[0]] = r[1];
  });
  return p;
}

function getSeasonality_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SEASONALITY);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  return data.map(function (r) {
    return { month_num: r[0], month_name: r[1], demand_index: r[2], fuel_price_index: r[3] };
  });
}

function getVehicleTiers_(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_VEHICLES);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  var tiers = {};
  data.forEach(function (r) {
    var tier = r[2];
    if (!tiers[tier]) {
      tiers[tier] = {
        count: 0, capacity_t_sum: 0, fuel_norm_sum: 0, maint_rate_sum: 0,
        driver_var_sum: 0, max_daily_km_sum: 0,
        fixed_cost_sum: 0
      };
    }
    var t = tiers[tier];
    t.count += 1;
    t.capacity_t_sum += r[3];
    t.fuel_norm_sum += r[4];
    t.fixed_cost_sum += (r[5] + r[6] + r[7]); // depreciation + insurance + driver fixed salary
    t.driver_var_sum += r[8];
    t.maint_rate_sum += r[9];
    t.max_daily_km_sum += r[10];
  });
  var workDays = params.work_days_per_month;
  Object.keys(tiers).forEach(function (k) {
    var t = tiers[k];
    t.avg_capacity_kg = (t.capacity_t_sum / t.count) * 1000;
    t.avg_fuel_norm = t.fuel_norm_sum / t.count;
    t.avg_maint_rate = t.maint_rate_sum / t.count;
    t.avg_driver_var_rate = t.driver_var_sum / t.count;
    t.avg_max_daily_km = t.max_daily_km_sum / t.count;
    t.capacity_km_month = t.count * workDays * t.avg_max_daily_km;
  });

  // Докупка машин (additional_vehicles_*): считаем, что новые машины того же вида транспорта
  // имеют характеристики, равные средним по уже имеющимся машинам этого вида. Влияет на
  // мощность парка и постоянные расходы (амортизация+страховка+оклад), но НЕ на среднюю
  // грузоподъёмность/расход топлива/тариф (они остаются средними по виду транспорта).
  var extraByTier = {
    Light: Number(params.additional_vehicles_light) || 0,
    Medium: Number(params.additional_vehicles_medium) || 0,
    Heavy: Number(params.additional_vehicles_heavy) || 0
  };
  Object.keys(tiers).forEach(function (k) {
    var extra = extraByTier[k] || 0;
    if (extra === 0) return;
    var t = tiers[k];
    var avgFixedPerVehicle = t.fixed_cost_sum / t.count;
    t.count = Math.max(0, t.count + extra);
    t.fixed_cost_sum = Math.max(0, t.fixed_cost_sum + extra * avgFixedPerVehicle);
    t.capacity_km_month = t.count * workDays * t.avg_max_daily_km;
  });
  return tiers;
}

function getCustomers_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CUSTOMERS);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  return data.map(function (r) {
    return {
      customer_id: r[0], customer_name: r[1], region: r[2], distance_km: r[3], segment: r[4],
      avg_order_weight_kg: r[5], orders_per_month: r[6], price_per_kg_rub: r[7],
      base_monthly_revenue_rub: r[8], tier: r[9]
    };
  });
}

// Тарифная сетка аутсорса: чем больше суммарный аутсорс-пробег компании за месяц, тем ниже
// ставка за км (скидка за объём в переговорах с перевозчиком). Пороги заданы в Params.
function getOutsourceRate_(totalOutsourceKm, params) {
  var t3 = Number(params.outsource_tier3_threshold_km_month);
  var t2 = Number(params.outsource_tier2_threshold_km_month);
  if (t3 && totalOutsourceKm >= t3) return Number(params.outsource_tier3_rate_rub_per_km);
  if (t2 && totalOutsourceKm >= t2) return Number(params.outsource_tier2_rate_rub_per_km);
  return Number(params.outsource_rate_rub_per_km);
}

// ---------------------------------------------------------------------------
// РАСЧЁТ БЮДЖЕТА ЗА 12 МЕСЯЦЕВ ДЛЯ ЗАДАННОГО СЦЕНАРИЯ (% ИЗМЕНЕНИЯ ВЫРУЧКИ)
// trackBreakdown=true дополнительно накапливает годовые итоги по видам транспорта и регионам.
// ---------------------------------------------------------------------------
// customers и params передаются вызывающей стороной (не читаются заново внутри), чтобы
// состав клиентов (после buildCustomerList_) был одинаковым во всех расчётах одного прогона
// (например, во всех 9 шагах сценарного анализа).
function computeMonthlyBudget_(scenarioPct, trackBreakdown, customers, params) {
  var seasonality = getSeasonality_();
  var tiers = getVehicleTiers_(params);

  var fixedMonth = 0;
  Object.keys(tiers).forEach(function (k) { fixedMonth += tiers[k].fixed_cost_sum; });
  fixedMonth += Number(params.overhead_monthly_fixed_rub);

  var scenarioMult = 1 + (Number(scenarioPct) / 100);
  var priceGrowthMult = 1 + (Number(params.price_growth_pct) || 0) / 100;
  var months = [];

  var tierAnnual = {}, regionAnnual = {}, bandAnnual = {};
  if (trackBreakdown) {
    Object.keys(tiers).forEach(function (k) {
      tierAnnual[k] = { revenue: 0, km: 0, fixed_cost: 0, fuel: 0, maintenance: 0, tolls: 0,
        driver_variable: 0, perdiem: 0, outsource: 0 };
    });
    DISTANCE_BANDS.forEach(function (b) { bandAnnual[b.label] = { revenue: 0, km: 0, cost: 0 }; });
  }

  seasonality.forEach(function (season) {
    var revenue = 0;
    var tierKm = {}, tierCat = {};
    Object.keys(tiers).forEach(function (k) {
      tierKm[k] = 0;
      tierCat[k] = { fuel: 0, maintenance: 0, tolls: 0, driver_variable: 0, perdiem: 0 };
    });
    var regionMonth = {}; // "region||tier" -> {region, tier, revenue, km, rawCost}
    var bandMonth = {}; // "band||tier" -> {band, tier, revenue, km, rawCost}

    customers.forEach(function (c) {
      var volumeKg = c.avg_order_weight_kg * c.orders_per_month * season.demand_index * scenarioMult;
      // Рост цены за кг увеличивает выручку, но НЕ объём/пробег/рейсы — поэтому не влияет
      // на транспортные расходы (чистый рост маржи, в отличие от роста объёма или числа клиентов).
      var custRevenue = volumeKg * c.price_per_kg_rub * priceGrowthMult;
      revenue += custRevenue;

      var t = tiers[c.tier];
      var capacityPerTripKg = t.avg_capacity_kg * Number(params.load_utilization);
      var trips = Math.max(1, Math.ceil(volumeKg / capacityPerTripKg));
      var roundTripKm = c.distance_km * 2;
      var totalKm = trips * roundTripKm;

      var fuelCost = totalKm * (t.avg_fuel_norm / 100) * Number(params.fuel_price_rub_per_l) * season.fuel_price_index;
      var maintCost = totalKm * t.avg_maint_rate;
      var tollKm = Math.max(0, c.distance_km - Number(params.toll_free_km)) * 2 * trips;
      var tollCost = tollKm * Number(params.toll_rate_rub_per_km);
      var driverVarCost = totalKm * t.avg_driver_var_rate;
      var tripDays = Math.ceil(roundTripKm / Number(params.avg_daily_driving_km));
      var nights = Math.max(0, tripDays - 1);
      var perdiemCost = nights * Number(params.perdiem_rate_rub_per_night) * trips;

      var tc = tierCat[c.tier];
      tc.fuel += fuelCost; tc.maintenance += maintCost; tc.tolls += tollCost;
      tc.driver_variable += driverVarCost; tc.perdiem += perdiemCost;
      tierKm[c.tier] += totalKm;

      if (trackBreakdown) {
        var rawCost = fuelCost + maintCost + tollCost + driverVarCost + perdiemCost;

        var key = c.region + '||' + c.tier;
        if (!regionMonth[key]) regionMonth[key] = { region: c.region, tier: c.tier, revenue: 0, km: 0, rawCost: 0 };
        var rm = regionMonth[key];
        rm.revenue += custRevenue;
        rm.km += totalKm;
        rm.rawCost += rawCost;

        var band = bandForDistance_(c.distance_km);
        var bKey = band + '||' + c.tier;
        if (!bandMonth[bKey]) bandMonth[bKey] = { band: band, tier: c.tier, revenue: 0, km: 0, rawCost: 0 };
        var bm = bandMonth[bKey];
        bm.revenue += custRevenue;
        bm.km += totalKm;
        bm.rawCost += rawCost;
      }
    });

    // Если требуемый пробег в тарифной группе превышает мощность собственного парка этой группы,
    // излишек уходит на аутсорс (3PL). Долю "своих" км считаем ОТДЕЛЬНО по каждому виду
    // транспорта (а не усреднённо по компании), чтобы верно относить затраты и на сам вид
    // транспорта, и на регионы/дистанции, которые он обслуживает. А вот ТАРИФ 3PL зависит от
    // СУММАРНОГО аутсорс-пробега компании за месяц (единый контракт с перевозчиком даёт скидку
    // за объём) — считаем его один раз на месяц и применяем ко всем видам транспорта одинаково.
    var ownShareByTier = {}, outsourceKmByTier = {}, totalOutsourceKm = 0;
    Object.keys(tiers).forEach(function (k) {
      var t = tiers[k];
      var planned = tierKm[k];
      var used = Math.min(planned, t.capacity_km_month);
      ownShareByTier[k] = planned > 0 ? (used / planned) : 1;
      var over = Math.max(0, planned - t.capacity_km_month);
      outsourceKmByTier[k] = over;
      totalOutsourceKm += over;
    });
    var effectiveOutsourceRate = getOutsourceRate_(totalOutsourceKm, params);
    var outsourceCostByTier = {}, outsourceCost = 0;
    Object.keys(tiers).forEach(function (k) {
      outsourceCostByTier[k] = outsourceKmByTier[k] * effectiveOutsourceRate;
      outsourceCost += outsourceCostByTier[k];
    });

    var catFuel = 0, catMaint = 0, catToll = 0, catDriverVar = 0, catPerdiem = 0;
    Object.keys(tiers).forEach(function (k) {
      var tc = tierCat[k];
      var share = ownShareByTier[k];
      tc.fuel *= share; tc.maintenance *= share; tc.tolls *= share;
      tc.driver_variable *= share; tc.perdiem *= share;
      catFuel += tc.fuel; catMaint += tc.maintenance; catToll += tc.tolls;
      catDriverVar += tc.driver_variable; catPerdiem += tc.perdiem;

      if (trackBreakdown) {
        var ta = tierAnnual[k];
        ta.km += tierKm[k];
        ta.fixed_cost += tiers[k].fixed_cost_sum;
        ta.fuel += tc.fuel; ta.maintenance += tc.maintenance; ta.tolls += tc.tolls;
        ta.driver_variable += tc.driver_variable; ta.perdiem += tc.perdiem;
        ta.outsource += outsourceCostByTier[k];
      }
    });

    if (trackBreakdown) {
      Object.keys(regionMonth).forEach(function (key) {
        var rm = regionMonth[key];
        var share = ownShareByTier[rm.tier];
        var finalCost = rm.rawCost * share + rm.km * (1 - share) * effectiveOutsourceRate;
        if (!regionAnnual[rm.region]) regionAnnual[rm.region] = { revenue: 0, km: 0, cost: 0 };
        regionAnnual[rm.region].revenue += rm.revenue;
        regionAnnual[rm.region].km += rm.km;
        regionAnnual[rm.region].cost += finalCost;
        tierAnnual[rm.tier].revenue += rm.revenue;
      });
      Object.keys(bandMonth).forEach(function (key) {
        var bm = bandMonth[key];
        var share = ownShareByTier[bm.tier];
        var finalCost = bm.rawCost * share + bm.km * (1 - share) * effectiveOutsourceRate;
        bandAnnual[bm.band].revenue += bm.revenue;
        bandAnnual[bm.band].km += bm.km;
        bandAnnual[bm.band].cost += finalCost;
      });
    }

    var variableTotal = catFuel + catMaint + catToll + catDriverVar + catPerdiem + outsourceCost;
    var totalCost = fixedMonth + variableTotal;

    months.push({
      month_num: season.month_num, month_name: season.month_name,
      revenue: revenue, fixed_cost: fixedMonth,
      fuel: catFuel, maintenance: catMaint, tolls: catToll,
      driver_variable: catDriverVar, perdiem: catPerdiem, outsource: outsourceCost,
      variable_total: variableTotal, total_cost: totalCost,
      cost_ratio_pct: revenue > 0 ? (totalCost / revenue) * 100 : 0
    });
  });

  var result = { months: months };
  if (trackBreakdown) {
    result.tierAnnual = tierAnnual;
    result.regionAnnual = regionAnnual;
    result.bandAnnual = bandAnnual;
  }
  return result;
}

// ---------------------------------------------------------------------------
// ЛИСТ "Бюджет"
// ---------------------------------------------------------------------------
function buildBudget() {
  var params = getParams_();
  var scenarioPct = Number(params.revenue_scenario_pct) || 0;
  var customers = buildCustomerList_(params);
  var months = computeMonthlyBudget_(scenarioPct, false, customers, params).months;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_BUDGET) || ss.insertSheet(SHEET_BUDGET);
  sh.clear();

  var header = ['Месяц', 'Выручка', 'Постоянные расходы', 'Топливо', 'ТО и ремонт',
    'Платные дороги', 'Доплата водителям', 'Суточные', 'Аутсорс (3PL)',
    'Переменные расходы, итого', 'Транспортные расходы, итого', 'Доля расходов в выручке, %'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  var rows = months.map(function (m) {
    return [m.month_name, m.revenue, m.fixed_cost, m.fuel, m.maintenance, m.tolls,
      m.driver_variable, m.perdiem, m.outsource, m.variable_total, m.total_cost, m.cost_ratio_pct];
  });
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  var totalsRowIdx = rows.length + 3; // строка 15 при 12 месяцах
  var sumCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  var totalsRow = ['ИТОГО ЗА ГОД'];
  sumCols.forEach(function (c) {
    var sum = rows.reduce(function (acc, r) { return acc + r[c - 1]; }, 0);
    totalsRow.push(sum);
  });
  var totalRevenue = totalsRow[1], totalCost = totalsRow[10];
  totalsRow.push(totalRevenue > 0 ? (totalCost / totalRevenue) * 100 : 0);
  sh.getRange(totalsRowIdx, 1, 1, totalsRow.length).setValues([totalsRow]).setFontWeight('bold');

  // Диапазон для круговой диаграммы (категории годовых расходов)
  var catValues = [
    totalsRow[2], // fixed
    totalsRow[3], // fuel
    totalsRow[4], // maintenance
    totalsRow[5], // tolls
    totalsRow[6], // driver variable
    totalsRow[7], // perdiem
    totalsRow[8]  // outsource
  ];
  sh.getRange(1, 14, 1, 2).setValues([['Категория', 'Расходы в год, руб']]).setFontWeight('bold');
  for (var i = 0; i < CATEGORY_LABELS.length; i++) {
    sh.getRange(2 + i, 14, 1, 2).setValues([[CATEGORY_LABELS[i], catValues[i]]]);
  }

  sh.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 3, rows.length, 8).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 12, rows.length, 1).setNumberFormat('0.0"%"');
  sh.autoResizeColumns(1, 15);
}

// ---------------------------------------------------------------------------
// ЛИСТ "Сценарии" — чувствительность затрат к изменению выручки
// ---------------------------------------------------------------------------
function runScenarioAnalysis() {
  var scenarioSteps = [-30, -20, -10, 0, 10, 20, 30, 40, 50];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SCENARIOS) || ss.insertSheet(SHEET_SCENARIOS);
  sh.clear();

  var header = ['Рост объёма у текущих клиентов, %', 'Выручка за год', 'Постоянные расходы за год',
    'Переменные расходы за год (вкл. аутсорс)', 'Транспортные расходы за год',
    'Доля расходов в выручке, %', 'Предельная стоимость доп. выручки, %'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  // Здесь варьируется только revenue_scenario_pct (объём у уже существующих клиентов).
  // customer_count_growth_pct и price_growth_pct берутся из Params и держатся постоянными
  // на протяжении всей развёртки — это отдельные, независимые рычаги роста (см. Params).
  var params = getParams_();
  var customers = buildCustomerList_(params);

  var results = [];
  scenarioSteps.forEach(function (pct) {
    var months = computeMonthlyBudget_(pct, false, customers, params).months;
    var revenue = months.reduce(function (a, m) { return a + m.revenue; }, 0);
    var fixed = months.reduce(function (a, m) { return a + m.fixed_cost; }, 0);
    var variable = months.reduce(function (a, m) { return a + m.variable_total; }, 0);
    var total = fixed + variable;
    results.push({ pct: pct, revenue: revenue, fixed: fixed, variable: variable, total: total });
  });

  var rows = results.map(function (r, i) {
    var marginal = '';
    if (i > 0) {
      var dRevenue = r.revenue - results[i - 1].revenue;
      var dCost = r.total - results[i - 1].total;
      marginal = dRevenue !== 0 ? (dCost / dRevenue) * 100 : 0;
    }
    return [r.pct, r.revenue, r.fixed, r.variable, r.total,
      r.revenue > 0 ? (r.total / r.revenue) * 100 : 0, marginal];
  });
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  sh.getRange(2, 2, rows.length, 4).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 6, rows.length, 2).setNumberFormat('0.0"%"');
  sh.autoResizeColumns(1, 7);
}

// ---------------------------------------------------------------------------
// ЛИСТЫ "По видам транспорта", "По регионам" и "Топ клиентов"
// ---------------------------------------------------------------------------
function buildTransportBreakdown() {
  var params = getParams_();
  var scenarioPct = Number(params.revenue_scenario_pct) || 0;
  var customers = buildCustomerList_(params);
  var result = computeMonthlyBudget_(scenarioPct, true, customers, params);
  var tiers = getVehicleTiers_(params);

  // Топ клиентов и сегменты ранжируются по базовой (без сезонности) выручке клиента —
  // сезонность и сценарий масштабируют всех одинаково, поэтому рейтинг от них не зависит.
  // Но чтобы годовые суммы на этих листах сходились с "Бюджетом", домножаем базовую выручку
  // на среднесезонный коэффициент и текущие сценарий/рост цены (тот же множитель для всех клиентов).
  var seasonality = getSeasonality_();
  var avgSeasonality = seasonality.reduce(function (s, se) { return s + se.demand_index; }, 0) / seasonality.length;
  var priceGrowthMult = 1 + (Number(params.price_growth_pct) || 0) / 100;
  var revenueScaleFactor = avgSeasonality * (1 + scenarioPct / 100) * priceGrowthMult;

  writeTierSheet_(result.tierAnnual, tiers, params);
  writeRegionSheet_(result.regionAnnual, customers);
  writeTopCustomersSheet_(customers, revenueScaleFactor);
  writeDistanceBandSheet_(result.bandAnnual, customers);
  writeSegmentSheet_(customers, revenueScaleFactor);
  writeCapexComparisonSheet_(params, scenarioPct, customers);
}

// Возвращает копию params с обнулёнными additional_vehicles_* — базовый сценарий "без докупки"
// для сравнения с текущими настройками пользователя (лист "CAPEX vs Аутсорс").
function paramsWithoutExtraVehicles_(params) {
  var clone = {};
  Object.keys(params).forEach(function (k) { clone[k] = params[k]; });
  clone.additional_vehicles_light = 0;
  clone.additional_vehicles_medium = 0;
  clone.additional_vehicles_heavy = 0;
  return clone;
}

// Сравнивает текущий парк (additional_vehicles_* = 0) с парком "докупили машины по Params" —
// показывает, окупается ли покупка новых машин экономией на аутсорсе. Спрос (customers) один
// и тот же в обоих сценариях — меняется только мощность/постоянные расходы парка.
function writeCapexComparisonSheet_(params, scenarioPct, customers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CAPEX) || ss.insertSheet(SHEET_CAPEX);
  sh.clear();

  var paramsBase = paramsWithoutExtraVehicles_(params);
  var tiersBase = getVehicleTiers_(paramsBase);
  var resultBase = computeMonthlyBudget_(scenarioPct, true, customers, paramsBase);

  var tiersBuy = getVehicleTiers_(params);
  var resultBuy = computeMonthlyBudget_(scenarioPct, true, customers, params);

  var extraByTier = {
    Light: Number(params.additional_vehicles_light) || 0,
    Medium: Number(params.additional_vehicles_medium) || 0,
    Heavy: Number(params.additional_vehicles_heavy) || 0
  };

  var header = ['Вид транспорта', 'Докупить машин (по Params)', 'Доп. постоянные расходы, руб/год',
    'Экономия на аутсорсе, руб/год', 'Чистый эффект за год (+ выгодно / − невыгодно)',
    'Загрузка парка без докупки, %', 'Загрузка парка с докупкой, %'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  var rows = TIER_ORDER.map(function (k) {
    var ab = resultBase.tierAnnual[k], au = resultBuy.tierAnnual[k];
    var extraFixed = au.fixed_cost - ab.fixed_cost;
    var outsourceSavings = ab.outsource - au.outsource; // положительное = экономия от докупки
    var netEffect = outsourceSavings - extraFixed;

    var capB = tiersBase[k].capacity_km_month * 12;
    var capU = tiersBuy[k].capacity_km_month * 12;
    var utilB = capB > 0 ? (ab.km / capB) * 100 : 0;
    var utilU = capU > 0 ? (au.km / capU) * 100 : 0;

    return [TIER_LABELS[k] || k, extraByTier[k] || 0, extraFixed, outsourceSavings, netEffect, utilB, utilU];
  });
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  var totalExtraFixed = rows.reduce(function (s, r) { return s + r[2]; }, 0);
  var totalSavings = rows.reduce(function (s, r) { return s + r[3]; }, 0);
  var totalNet = rows.reduce(function (s, r) { return s + r[4]; }, 0);
  sh.getRange(rows.length + 2, 1, 1, header.length).setValues([[
    'ИТОГО', extraByTier.Light + extraByTier.Medium + extraByTier.Heavy,
    totalExtraFixed, totalSavings, totalNet, '', ''
  ]]).setFontWeight('bold');

  sh.getRange(2, 3, rows.length, 3).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 6, rows.length, 2).setNumberFormat('0.0"%"');
  sh.getRange(rows.length + 2, 3, 1, 3).setNumberFormat('#,##0 ₽');

  sh.getRange(rows.length + 4, 1).setValue(
    'Примечание: "Экономия на аутсорсе" может быть не 0 даже там, где докупки нет — тариф 3PL общий ' +
    'на всю компанию (см. outsource_tier2/3_* в Params) и зависит от суммарного аутсорс-пробега; ' +
    'докупка машин в одном виде транспорта может сдвинуть компанию между тарифными порогами и для остальных.'
  );

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

function writeTopCustomersSheet_(customers, revenueScaleFactor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_TOP_CUSTOMERS) || ss.insertSheet(SHEET_TOP_CUSTOMERS);
  sh.clear();

  var header = ['Клиент', 'Регион', 'Сегмент', 'Вид транспорта', 'Расстояние, км',
    'Выручка за год', 'Доля в общей выручке, %'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  // Ранжируем по базовой выручке клиента — сезонность и сценарий масштабируют всех клиентов
  // одинаково, поэтому относительный рейтинг от них не зависит. revenueScaleFactor приводит
  // отображаемые суммы к текущему сценарию, чтобы годовые итоги сходились с листом "Бюджет".
  var totalRevenue = customers.reduce(function (s, c) { return s + c.base_monthly_revenue_rub * 12 * revenueScaleFactor; }, 0);
  var sorted = customers.slice().sort(function (a, b) { return b.base_monthly_revenue_rub - a.base_monthly_revenue_rub; });

  var rows = sorted.slice(0, TOP_CUSTOMERS_N).map(function (c) {
    var annualRevenue = c.base_monthly_revenue_rub * 12 * revenueScaleFactor;
    return [c.customer_name, c.region, c.segment, TIER_LABELS[c.tier] || c.tier, c.distance_km,
      annualRevenue, (annualRevenue / totalRevenue) * 100];
  });
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  sh.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('0.00"%"');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

function writeTierSheet_(tierAnnual, tiers, params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_TRANSPORT_TYPES) || ss.insertSheet(SHEET_TRANSPORT_TYPES);
  sh.clear();

  var header = ['Вид транспорта', 'Кол-во машин', 'Выручка за год', 'Пробег (план), км/год',
    'Мощность парка, км/год', 'Загрузка парка, %', 'Постоянные расходы',
    'Переменные расходы (свой парк)', 'Аутсорс (3PL)', 'Итого расходы', 'Доля расходов в выручке, %',
    'Себестоимость своего парка, руб/км', 'Ставка аутсорса (факт.), руб/км', 'Аутсорс дороже, раз'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  var baseOutsourceRate = Number(params.outsource_rate_rub_per_km);
  var rows = TIER_ORDER.filter(function (k) { return tierAnnual[k]; }).map(function (k) {
    var ta = tierAnnual[k];
    var t = tiers[k];
    var capacityAnnualKm = t.capacity_km_month * 12;
    var utilizationPct = capacityAnnualKm > 0 ? (ta.km / capacityAnnualKm) * 100 : 0;
    var ownVariable = ta.fuel + ta.maintenance + ta.tolls + ta.driver_variable + ta.perdiem;
    var totalCost = ta.fixed_cost + ownVariable + ta.outsource;
    var costRatio = ta.revenue > 0 ? (totalCost / ta.revenue) * 100 : 0;
    var ownKmUsed = Math.min(ta.km, capacityAnnualKm);
    var ownCostPerKm = ownKmUsed > 0 ? ownVariable / ownKmUsed : 0;
    // Фактическая ставка аутсорса — может быть ниже базовой, если сработала скидка за объём
    // (см. outsource_tier2/3_* в Params): считаем её из факта (расходы / аутсорс-км).
    var outsourceKmTier = Math.max(0, ta.km - ownKmUsed);
    var effectiveOutsourceRate = outsourceKmTier > 0 ? ta.outsource / outsourceKmTier : baseOutsourceRate;
    var outsourceMultiplier = ownCostPerKm > 0 ? effectiveOutsourceRate / ownCostPerKm : 0;
    return [TIER_LABELS[k] || k, t.count, ta.revenue, ta.km, capacityAnnualKm, utilizationPct,
      ta.fixed_cost, ownVariable, ta.outsource, totalCost, costRatio,
      ownCostPerKm, effectiveOutsourceRate, outsourceMultiplier];
  });
  var nTiers = rows.length;
  sh.getRange(2, 1, nTiers, header.length).setValues(rows);

  // Общие накладные расходы не относятся ни к одному виду транспорта — показываем отдельно,
  // чтобы строка "ИТОГО" сходилась с итогом на листе "Бюджет".
  var overheadAnnual = Number(params.overhead_monthly_fixed_rub) * 12;
  var overheadRowIdx = nTiers + 3;
  sh.getRange(overheadRowIdx, 1, 1, header.length).setValues([[
    'Общие накладные расходы (не относятся к виду транспорта)', '', '', '', '', '',
    overheadAnnual, '', '', overheadAnnual, '', '', '', ''
  ]]);

  var totalVehicles = rows.reduce(function (s, r) { return s + r[1]; }, 0);
  var totalRevenue = rows.reduce(function (s, r) { return s + r[2]; }, 0);
  var totalKm = rows.reduce(function (s, r) { return s + r[3]; }, 0);
  var totalCapacity = rows.reduce(function (s, r) { return s + r[4]; }, 0);
  var totalUtil = totalCapacity > 0 ? (totalKm / totalCapacity) * 100 : 0;
  var totalFixed = rows.reduce(function (s, r) { return s + r[6]; }, 0) + overheadAnnual;
  var totalOwnVar = rows.reduce(function (s, r) { return s + r[7]; }, 0);
  var totalOutsource = rows.reduce(function (s, r) { return s + r[8]; }, 0);
  var totalCost = totalFixed + totalOwnVar + totalOutsource;
  var totalRatio = totalRevenue > 0 ? (totalCost / totalRevenue) * 100 : 0;
  var totalOwnKmUsed = Math.min(totalKm, totalCapacity);
  var totalOwnCostPerKm = totalOwnKmUsed > 0 ? totalOwnVar / totalOwnKmUsed : 0;
  var totalOutsourceKmAll = Math.max(0, totalKm - totalOwnKmUsed);
  var totalEffectiveOutsourceRate = totalOutsourceKmAll > 0 ? totalOutsource / totalOutsourceKmAll : baseOutsourceRate;
  var totalMultiplier = totalOwnCostPerKm > 0 ? totalEffectiveOutsourceRate / totalOwnCostPerKm : 0;
  sh.getRange(overheadRowIdx + 1, 1, 1, header.length).setValues([[
    'ИТОГО', totalVehicles, totalRevenue, totalKm, totalCapacity, totalUtil,
    totalFixed, totalOwnVar, totalOutsource, totalCost, totalRatio,
    totalOwnCostPerKm, totalEffectiveOutsourceRate, totalMultiplier
  ]]).setFontWeight('bold');

  sh.getRange(2, 3, nTiers, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 7, nTiers, 4).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 6, nTiers, 1).setNumberFormat('0.0"%"');
  sh.getRange(2, 11, nTiers, 1).setNumberFormat('0.0"%"');
  sh.getRange(2, 12, nTiers, 2).setNumberFormat('0.0 "₽/км"');
  sh.getRange(2, 14, nTiers, 1).setNumberFormat('0.00"×"');
  sh.getRange(overheadRowIdx, 7, 2, 4).setNumberFormat('#,##0 ₽');
  sh.getRange(overheadRowIdx + 1, 12, 1, 2).setNumberFormat('0.0 "₽/км"');
  sh.getRange(overheadRowIdx + 1, 14, 1, 1).setNumberFormat('0.00"×"');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

function writeRegionSheet_(regionAnnual, customers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_REGIONS) || ss.insertSheet(SHEET_REGIONS);
  sh.clear();

  var header = ['Регион', 'Кол-во клиентов', 'Выручка за год', 'Пробег, км/год',
    'Переменные расходы за год', 'Доля переменных расходов в выручке, %', 'Средняя дистанция, км'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  var custStats = {};
  customers.forEach(function (c) {
    if (!custStats[c.region]) custStats[c.region] = { count: 0, distSum: 0 };
    custStats[c.region].count += 1;
    custStats[c.region].distSum += c.distance_km;
  });

  var rows = Object.keys(regionAnnual).map(function (region) {
    var ra = regionAnnual[region];
    var cs = custStats[region] || { count: 0, distSum: 0 };
    var costRatio = ra.revenue > 0 ? (ra.cost / ra.revenue) * 100 : 0;
    var avgDistance = cs.count > 0 ? cs.distSum / cs.count : 0;
    return [region, cs.count, ra.revenue, ra.km, ra.cost, costRatio, avgDistance];
  });
  rows.sort(function (a, b) { return b[2] - a[2]; }); // по убыванию выручки
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  sh.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('0.0"%"');
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('0 "км"');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

// ---------------------------------------------------------------------------
// ЛИСТ "По дистанции" — как расходы и их доля в выручке зависят от расстояния до клиента
// ---------------------------------------------------------------------------
function writeDistanceBandSheet_(bandAnnual, customers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_DISTANCE_BANDS) || ss.insertSheet(SHEET_DISTANCE_BANDS);
  sh.clear();

  var header = ['Дистанция', 'Кол-во клиентов', 'Выручка за год', 'Пробег, км/год',
    'Переменные расходы за год', 'Доля переменных расходов в выручке, %', 'Себестоимость, руб/км'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  var custCount = {};
  DISTANCE_BANDS.forEach(function (b) { custCount[b.label] = 0; });
  customers.forEach(function (c) {
    var band = bandForDistance_(c.distance_km);
    custCount[band] += 1;
  });

  var rows = DISTANCE_BANDS.map(function (b) {
    var ba = bandAnnual[b.label] || { revenue: 0, km: 0, cost: 0 };
    var costRatio = ba.revenue > 0 ? (ba.cost / ba.revenue) * 100 : 0;
    var costPerKm = ba.km > 0 ? ba.cost / ba.km : 0;
    return [b.label, custCount[b.label], ba.revenue, ba.km, ba.cost, costRatio, costPerKm];
  });
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  sh.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('0.0"%"');
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('0.0 "₽/км"');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

// ---------------------------------------------------------------------------
// ЛИСТ "По сегментам клиентов" — экономика клиента по типу (розница/опт/крупный дистрибьютор)
// ---------------------------------------------------------------------------
function writeSegmentSheet_(customers, revenueScaleFactor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SEGMENTS) || ss.insertSheet(SHEET_SEGMENTS);
  sh.clear();

  var header = ['Сегмент', 'Кол-во клиентов', 'Выручка за год', 'Доля в выручке, %',
    'Выручка на клиента в год', 'Средний вес заказа, кг', 'Заказов в месяц', 'Средняя дистанция, км'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');

  var groups = {};
  customers.forEach(function (c) {
    if (!groups[c.segment]) groups[c.segment] = { count: 0, revenue: 0, weightSum: 0, ordersSum: 0, distSum: 0 };
    var g = groups[c.segment];
    var annualRevenue = c.base_monthly_revenue_rub * 12 * revenueScaleFactor;
    g.count += 1;
    g.revenue += annualRevenue;
    g.weightSum += c.avg_order_weight_kg;
    g.ordersSum += c.orders_per_month;
    g.distSum += c.distance_km;
  });

  var totalRevenue = Object.keys(groups).reduce(function (s, k) { return s + groups[k].revenue; }, 0);
  var rows = SEGMENT_ORDER.filter(function (k) { return groups[k]; }).map(function (k) {
    var g = groups[k];
    return [k, g.count, g.revenue, totalRevenue > 0 ? (g.revenue / totalRevenue) * 100 : 0,
      g.revenue / g.count, g.weightSum / g.count, g.ordersSum / g.count, g.distSum / g.count];
  });
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  sh.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 4, rows.length, 1).setNumberFormat('0.0"%"');
  sh.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 8, rows.length, 1).setNumberFormat('0 "км"');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

// ---------------------------------------------------------------------------
// ЛИСТ "Дашборд" — графики
// ---------------------------------------------------------------------------
function buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var budgetSh = ss.getSheetByName(SHEET_BUDGET);
  var scenariosSh = ss.getSheetByName(SHEET_SCENARIOS);
  var tierSh = ss.getSheetByName(SHEET_TRANSPORT_TYPES);
  var regionSh = ss.getSheetByName(SHEET_REGIONS);
  var topCustSh = ss.getSheetByName(SHEET_TOP_CUSTOMERS);
  var bandSh = ss.getSheetByName(SHEET_DISTANCE_BANDS);
  var segmentSh = ss.getSheetByName(SHEET_SEGMENTS);
  var capexSh = ss.getSheetByName(SHEET_CAPEX);
  if (!budgetSh || !scenariosSh || !tierSh || !regionSh || !topCustSh || !bandSh || !segmentSh || !capexSh) {
    SpreadsheetApp.getUi().alert('Сначала выполните шаги 1-3: "Пересчитать бюджет", "Сценарный анализ выручки", "Разбивка по видам транспорта, регионам, клиентам, дистанции и сегментам".');
    return;
  }

  var sh = ss.getSheetByName(SHEET_DASHBOARD) || ss.insertSheet(SHEET_DASHBOARD);
  sh.clear();
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.getRange(1, 1).setValue('ДАШБОРД: ТРАНСПОРТНЫЙ БЮДЖЕТ').setFontWeight('bold').setFontSize(14);

  var nMonthRows = 12;

  // 1. Выручка vs транспортные расходы по месяцам
  var chart1 = budgetSh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(budgetSh.getRange(1, 1, nMonthRows + 1, 1))   // месяцы
    .addRange(budgetSh.getRange(1, 2, nMonthRows + 1, 1))   // выручка
    .addRange(budgetSh.getRange(1, 11, nMonthRows + 1, 1))  // итого расходы
    .setOption('title', 'Выручка и транспортные расходы по месяцам')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'top' })
    .setPosition(3, 1, 0, 0)
    .build();
  sh.insertChart(chart1);

  // 2. Постоянные vs переменные расходы (structure) по месяцам
  var chart2 = budgetSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(budgetSh.getRange(1, 1, nMonthRows + 1, 1))   // месяцы
    .addRange(budgetSh.getRange(1, 3, nMonthRows + 1, 1))   // постоянные
    .addRange(budgetSh.getRange(1, 10, nMonthRows + 1, 1))  // переменные итого
    .setOption('title', 'Структура расходов: постоянные vs переменные')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', true)
    .setOption('legend', { position: 'top' })
    .setPosition(3, 8, 0, 0)
    .build();
  sh.insertChart(chart2);

  // 3. Эластичность: доля затрат в выручке при разных сценариях изменения выручки
  var nScenarioRows = scenariosSh.getLastRow() - 1;
  var chart3 = scenariosSh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(scenariosSh.getRange(1, 1, nScenarioRows + 1, 1))  // % изменения выручки
    .addRange(scenariosSh.getRange(1, 6, nScenarioRows + 1, 1))  // доля расходов в выручке
    .setOption('title', 'Как меняется доля транспортных расходов в выручке при изменении выручки')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('hAxis', { title: 'Изменение выручки, %' })
    .setOption('vAxis', { title: 'Расходы / выручка, %' })
    .setOption('legend', { position: 'none' })
    .setPosition(20, 1, 0, 0)
    .build();
  sh.insertChart(chart3);

  // 4. Круговая диаграмма — структура годовых расходов по статьям
  var chart4 = budgetSh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(budgetSh.getRange(1, 14, CATEGORY_LABELS.length + 1, 2))
    .setOption('title', 'Структура транспортных расходов за год по статьям')
    .setOption('pieHole', 0.4)
    .setPosition(20, 8, 0, 0)
    .build();
  sh.insertChart(chart4);

  // 5. Структура расходов по видам транспорта: постоянные / переменные (свой парк) / аутсорс
  var nTierRows = TIER_ORDER.length;
  var chart5 = tierSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(tierSh.getRange(1, 1, nTierRows + 1, 1))  // вид транспорта
    .addRange(tierSh.getRange(1, 7, nTierRows + 1, 1))  // постоянные
    .addRange(tierSh.getRange(1, 8, nTierRows + 1, 1))  // переменные (свой парк)
    .addRange(tierSh.getRange(1, 9, nTierRows + 1, 1))  // аутсорс
    .setOption('title', 'Структура расходов по видам транспорта')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('isStacked', true)
    .setOption('legend', { position: 'top' })
    .setPosition(37, 1, 0, 0)
    .build();
  sh.insertChart(chart5);

  // 6. Загрузка собственного парка по видам транспорта, %
  var chart6 = tierSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(tierSh.getRange(1, 1, nTierRows + 1, 1))  // вид транспорта
    .addRange(tierSh.getRange(1, 6, nTierRows + 1, 1))  // загрузка парка, %
    .setOption('title', 'Загрузка собственного парка по видам транспорта, %')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setPosition(37, 8, 0, 0)
    .build();
  sh.insertChart(chart6);

  // 7. Выручка по регионам
  var nRegionRows = regionSh.getLastRow() - 1;
  var chart7 = regionSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(regionSh.getRange(1, 1, nRegionRows + 1, 1))  // регион
    .addRange(regionSh.getRange(1, 3, nRegionRows + 1, 1))  // выручка за год
    .setOption('title', 'Выручка по регионам за год')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setPosition(54, 1, 0, 0)
    .build();
  sh.insertChart(chart7);

  // 8. Доля переменных транспортных расходов в выручке по регионам
  var chart8 = regionSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(regionSh.getRange(1, 1, nRegionRows + 1, 1))  // регион
    .addRange(regionSh.getRange(1, 6, nRegionRows + 1, 1))  // доля переменных расходов в выручке, %
    .setOption('title', 'Доля переменных транспортных расходов в выручке по регионам, %')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setPosition(54, 8, 0, 0)
    .build();
  sh.insertChart(chart8);

  // 9. Клиенты с наибольшей выручкой
  var nTopCustRows = topCustSh.getLastRow() - 1;
  var chart9 = topCustSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(topCustSh.getRange(1, 1, nTopCustRows + 1, 1))  // клиент
    .addRange(topCustSh.getRange(1, 6, nTopCustRows + 1, 1))  // выручка за год
    .setOption('title', 'Топ-' + nTopCustRows + ' клиентов по выручке за год')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setPosition(71, 1, 0, 0)
    .build();
  sh.insertChart(chart9);

  // 10. Выручка по дистанционным поясам
  var nBandRows = bandSh.getLastRow() - 1;
  var chart10 = bandSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(bandSh.getRange(1, 1, nBandRows + 1, 1))  // дистанция
    .addRange(bandSh.getRange(1, 3, nBandRows + 1, 1))  // выручка за год
    .setOption('title', 'Выручка по дистанционным поясам')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setPosition(88, 1, 0, 0)
    .build();
  sh.insertChart(chart10);

  // 11. Доля переменных расходов в выручке по дистанционным поясам
  var chart11 = bandSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(bandSh.getRange(1, 1, nBandRows + 1, 1))  // дистанция
    .addRange(bandSh.getRange(1, 6, nBandRows + 1, 1))  // доля переменных расходов в выручке, %
    .setOption('title', 'Доля переменных расходов в выручке по дистанционным поясам')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setPosition(88, 8, 0, 0)
    .build();
  sh.insertChart(chart11);

  // 12. Выручка на клиента по сегментам
  var nSegmentRows = segmentSh.getLastRow() - 1;
  var chart12 = segmentSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(segmentSh.getRange(1, 1, nSegmentRows + 1, 1))  // сегмент
    .addRange(segmentSh.getRange(1, 5, nSegmentRows + 1, 1))  // выручка на клиента в год
    .setOption('title', 'Выручка на клиента в год по сегментам')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setPosition(105, 1, 0, 0)
    .build();
  sh.insertChart(chart12);

  // 13. Себестоимость км: свой парк vs аутсорс, по видам транспорта
  var chart13 = tierSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(tierSh.getRange(1, 1, nTierRows + 1, 1))   // вид транспорта
    .addRange(tierSh.getRange(1, 12, nTierRows + 1, 1))  // себестоимость своего парка, руб/км
    .addRange(tierSh.getRange(1, 13, nTierRows + 1, 1))  // ставка аутсорса, руб/км
    .setOption('title', 'Себестоимость км: свой парк vs аутсорс')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'top' })
    .setPosition(105, 8, 0, 0)
    .build();
  sh.insertChart(chart13);

  // 14. CAPEX vs аутсорс: чистый эффект от докупки машин по видам транспорта
  var nCapexRows = TIER_ORDER.length;
  var chart14 = capexSh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(capexSh.getRange(1, 1, nCapexRows + 1, 1))  // вид транспорта
    .addRange(capexSh.getRange(1, 5, nCapexRows + 1, 1))  // чистый эффект за год
    .setOption('title', 'Докупка машин: чистый эффект за год (+ выгодно / − невыгодно)')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('legend', { position: 'none' })
    .setPosition(122, 1, 0, 0)
    .build();
  sh.insertChart(chart14);

  sh.setColumnWidth(1, 120);
}
