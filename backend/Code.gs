// Kamyl App - приём замеров къамылей в Google Таблицу.
// Вставьте этот код в Расширения > Apps Script своей таблицы и опубликуйте как веб-приложение.

var SHEET_NAME = 'data';
var HEAD = ['Время', 'Статус', 'Версия', 'Язык', 'Имя / Instagram', 'Комментарий',
  'Длина', 'Внутр. диам.', 'Внешн. диам.', 'Диам. отв.', 'Отв. 1', 'Отв. 2', 'Отв. 3', 'Темп.',
  'Тоника', 'Тоника, ц',
  'Гц р1 закр.', 'Гц р1 о1', 'Гц р1 о1-2', 'Гц р1 о1-3',
  'Гц р2 закр.', 'Гц р2 о1', 'Гц р2 о1-2', 'Гц р2 о1-3',
  'Разброс, ц (8)', 'Откл. тоника, ц', 'Откл. о1, ц', 'Откл. о1-2, ц', 'Откл. о1-3, ц', 'ID устройства'];

// те же допуски и формулы, что в приложении
var BLOW = 55.5, ABS_TOL = 100, INT_TOL = 50, OCT_TOL = 60, SD_MAX = 12, PER_HOUR = 5;

function doGet() { return out_({ ok: true, service: 'kamyl-app' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents);
    if (d.hp) return out_({ ok: true });                       // ловушка для ботов
    var err = check_(d);
    if (err) return out_({ ok: false, error: err });
    var cache = CacheService.getScriptCache(), key = 'n_' + String(d.sid || 'x').slice(0, 20);
    var n = Number(cache.get(key) || 0);
    if (n >= PER_HOUR) return out_({ ok: false, error: 'too many submissions, try later' });
    cache.put(key, String(n + 1), 3600);
    var dev = model_(d);
    sheet_().appendRow([new Date(), 'new', str_(d.v, 10), str_(d.lang, 5), str_(d.who, 80), str_(d.note, 300),
      d.L, d.bore, d.outer, d.hole, d.e1, d.e2, d.e3, d.temp === null ? '' : d.temp,
      str_(d.tonic, 5), d.tonicC].concat(d.f).concat([d.sd.join(' ')]).concat(dev.map(Math.round)).concat([str_(d.sid, 20)]));
    return out_({ ok: true });
  } catch (x) {
    return out_({ ok: false, error: 'bad request' });
  } finally {
    try { lock.releaseLock(); } catch (y) {}
  }
}

function check_(d) {
  var num = function (v, a, b) { return typeof v === 'number' && isFinite(v) && v >= a && v <= b; };
  if (!num(d.L, 200, 1200) || !num(d.bore, 5, 40) || !num(d.outer, d.bore, d.bore + 20) || d.outer <= d.bore) return 'tube size';
  if (!num(d.hole, 2, 20)) return 'hole size';
  if (!(num(d.e3, 0, d.L) && d.e3 > 0 && d.e2 > d.e3 && d.e1 > d.e2 && d.e1 + d.hole < d.L)) return 'hole order';
  if (d.e1 - d.e2 < d.hole || d.e2 - d.e3 < d.hole) return 'hole gap';
  if (d.temp !== null && !num(d.temp, -10, 50)) return 'temperature';
  if (!(d.f instanceof Array) || d.f.length !== 8 || !d.f.every(function (x) { return num(x, 60, 2500); })) return 'notes';
  if (!(d.sd instanceof Array) || d.sd.length !== 8 || !d.sd.every(function (x) { return num(x, 0, SD_MAX); })) return 'unsteady notes';
  for (var i = 1; i < 4; i++) if (!(d.f[i] > d.f[i - 1] && d.f[i + 4] > d.f[i + 3])) return 'note order';
  for (i = 0; i < 4; i++) if (Math.abs(1200 * log2_(d.f[i + 4] / d.f[i]) - 1200) > OCT_TOL) return 'octaves';
  var dev = model_(d);
  if (Math.abs(dev[0]) > ABS_TOL) return 'tube length does not match the sound';
  for (i = 1; i < 4; i++) if (Math.abs(dev[i]) > INT_TOL) return 'hole ' + i + ' does not match the sound';
  return '';
}

// отклонение сыгранного от расчёта: тоника - по высоте, остальные - по интервалу от тоники
function model_(d) {
  var c = 331.3 + 0.606 * (d.temp === null ? 25 : d.temp);
  var f = function (Le) { return c * 1000 / (2 * Le); };
  var te = (d.outer - d.bore) / 2 + 0.75 * d.hole, R = Math.pow(d.bore / d.hole, 2);
  var mf = [f(d.L + BLOW + 0.6133 * d.bore / 2)], prev = d.L;
  [d.e1, d.e2, d.e3].forEach(function (e) {
    var x = e + d.hole / 2, g = Math.max(prev - x, 1e-6), z = g / 2 * (Math.sqrt(1 + 4 * te / g * R) - 1);
    mf.push(f(x + z + BLOW)); prev = x;
  });
  var dev = [1200 * log2_(d.f[0] / mf[0])];
  for (var k = 1; k < 4; k++) dev.push(1200 * log2_(d.f[k] / d.f[0]) - 1200 * log2_(mf[k] / mf[0]));
  return dev;
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) { sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function log2_(x) { return Math.log(x) / Math.LN2; }
function str_(v, n) { return String(v === undefined || v === null ? '' : v).slice(0, n).replace(/^[=+\-@]/, "'$&"); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Запустите один раз вручную, чтобы создать лист с заголовками
function setup() { sheet_(); }
