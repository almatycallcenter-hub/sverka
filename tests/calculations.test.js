(function(){
  try {
  var calculations;
  if (typeof WScript !== 'undefined'){
    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var modulePath = fso.BuildPath(
      fso.GetParentFolderName(WScript.ScriptFullName), '..\\calculations.js');
    if (!Array.prototype.forEach) Array.prototype.forEach = function(fn){
      for (var i = 0; i < this.length; i++) if (i in this) fn(this[i], i, this);
    };
    if (!Array.prototype.filter) Array.prototype.filter = function(fn){
      var out = [];
      this.forEach(function(value, i, array){ if (fn(value, i, array)) out.push(value); });
      return out;
    };
    if (!Array.prototype.map) Array.prototype.map = function(fn){
      var out = [];
      this.forEach(function(value, i, array){ out.push(fn(value, i, array)); });
      return out;
    };
    if (!String.prototype.trim) String.prototype.trim = function(){
      return this.replace(/^\s+|\s+$/g, '');
    };
    eval(fso.OpenTextFile(modulePath, 1).ReadAll());
    calculations = this.Calculations;
  } else {
    calculations = require('../calculations.js');
  }

  var passed = 0;
  function equal(actual, expected, label){
    if (actual !== expected)
      throw new Error(label + ': expected ' + expected + ', got ' + actual);
    passed++;
  }
  function close(actual, expected, label){
    if (Math.abs(actual - expected) > 1e-9)
      throw new Error(label + ': expected ' + expected + ', got ' + actual);
    passed++;
  }
  function daysBetween(a, b){
    function date(value){
      var parts = value.split('-');
      return new Date(+parts[0], +parts[1] - 1, +parts[2]);
    }
    return (date(b) - date(a)) / 86400000;
  }

  var fifo = calculations.fifoAges([
    { d:'2026-01-01', p:3, s:2, l:1 },
    { d:'2026-01-02', p:2, s:3, l:0 },
    { d:'2026-01-03', p:0, s:1, l:0 }
  ], 2, false, daysBetween);
  equal(fifo.ages[0], null, 'opening stock has unknown age');
  close(fifo.ages[1], 2 / 3, 'FIFO age includes older remaining stock');
  equal(fifo.ages[2], 1, 'FIFO age rolls across days');
  close(fifo.average, 0.75, 'average excludes opening stock and losses');

  var delivery = calculations.fifoAges([
    { d:'2026-02-01', p:2, s:0, l:0, dost:1 },
    { d:'2026-02-02', p:1, s:2, l:0, dost:0 }
  ], 0, false, daysBetween);
  close(delivery.ages[1], 0.5, 'delivery removes stock from the FIFO queue');

  var frozen = calculations.fifoAges([
    { d:'2026-01-01', p:2, s:1, l:0 }
  ], 0, true, daysBetween);
  equal(frozen.ages[0], null, 'frozen product has no tracked age');
  equal(frozen.average, null, 'frozen product has no average age');

  equal(calculations.recommendedOrder([
    { s:2, st:0 }, { s:4, st:0 }, { s:100, st:1 }
  ]), 4, 'recommended order uses non-stockout median plus 10%');
  equal(calculations.recommendedOrder([{ s:0, st:0 }, { s:5, st:1 }]), 0,
    'no available demand means no recommended order');

  equal(calculations.reconcileDiscount(5, 5, false, false), 'ок',
    'matching listed quantity');
  equal(calculations.reconcileDiscount(5, 5.0000005, false, false), 'ок',
    'reconciliation tolerance');
  equal(calculations.reconcileDiscount(5, 6, false, false), 'продажа больше списка',
    'sales exceed the list');
  equal(calculations.reconcileDiscount(5, 4, false, false), 'продажа меньше списка',
    'sales below the list');
  equal(calculations.reconcileDiscount(0, 0, false, false), 'ок',
    'empty list and no sales');
  equal(calculations.reconcileDiscount(0, 1, false, false),
    'нарушение (нет в списке филиала)', 'sales missing from the list');
  equal(calculations.reconcileDiscount(5, 0, true, false), 'нет данных',
    'blank listed quantity is not compared');
  equal(calculations.reconcileDiscount(5, 0, false, true), 'нет данных',
    'dates beyond the last list are not compared');

  close(calculations.discountPercent('Скидка 10%'), 0.1, 'discount percentage');
  close(calculations.discountPercent('50% сотрудникам'), 0.5, 'discount percentage placement');
  equal(calculations.discountPercent('Бонусы'), null, 'non-percentage discount');
  close(calculations.estimatedDiscountAmount(4, 1 / 8, 1000, 'Скидка 20%'), 100,
    'discount estimate accounts for piece size');
  equal(calculations.estimatedDiscountAmount(4, 1, 1000, 'Бонусы'), null,
    'cannot estimate amount without a percentage');

  if (typeof WScript !== 'undefined') WScript.Echo('Passed ' + passed + ' checks.');
  else console.log('Passed ' + passed + ' checks.');
  } catch(error) {
    if (typeof WScript !== 'undefined'){
      WScript.Echo(error.message || String(error));
      WScript.Quit(1);
    } else {
      console.error(error);
      process.exitCode = 1;
    }
  }
})();
