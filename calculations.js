/* Чистые расчёты сайта сверки — без страницы, без данных, без браузера.
   Их проверяет tests/calculations.test.js (node tests/calculations.test.js).
   Вынесено GitHub Copilot 13.10; колонки отчёта живут в index.html (REP_COLS). */
(function(root){

  var calculations = {
    fifoAges: function(series, open, frozenPos, daysBetween){
      var fifo = [], ageSum = 0, ageQty = 0, ages = [];
      if (frozenPos) {
        for (var i = 0; i < series.length; i++) ages.push(null);
        return { ages: ages, average: null };
      }
      if (open > 1e-9) fifo.push({ d: '', qty: open });
      series.forEach(function(x){
        if (x.p > 1e-9) fifo.push({ d: x.d, qty: x.p });
        var need = x.s, dSum = 0, dQty = 0;
        while (need > 1e-9 && fifo.length){
          var batch = fifo[0], taken = Math.min(batch.qty, need);
          if (batch.d){
            var age = daysBetween(batch.d, x.d);
            ageSum += age * taken;
            ageQty += taken;
            dSum += age * taken;
            dQty += taken;
          }
          batch.qty -= taken;
          need -= taken;
          if (batch.qty <= 1e-9) fifo.shift();
        }
        var losses = x.l + (x.dost || 0);
        while (losses > 1e-9 && fifo.length){
          var lossBatch = fifo[0], lost = Math.min(lossBatch.qty, losses);
          lossBatch.qty -= lost;
          losses -= lost;
          if (lossBatch.qty <= 1e-9) fifo.shift();
        }
        ages.push(dQty > 0 ? dSum / dQty : null);
      });
      return { ages: ages, average: ageQty > 0 ? ageSum / ageQty : null };
    },

    recommendedOrder: function(series){
      var demand = series.filter(function(x){ return !x.st; }).map(function(x){ return x.s; })
                         .sort(function(a, b){ return a - b; });
      var median = demand.length
        ? (demand.length % 2 ? demand[(demand.length - 1) / 2]
           : (demand[demand.length / 2 - 1] + demand[demand.length / 2]) / 2)
        : 0;
      return median > 0 ? Math.ceil(median * 1.1) : 0;
    },

    reconcileDiscount: function(listed, sold, noQty, afterLast){
      if (noQty || afterLast) return 'нет данных';
      if (listed > 0){
        return Math.abs(sold - listed) < 1e-6 ? 'ок'
          : (sold > listed ? 'продажа больше списка' : 'продажа меньше списка');
      }
      return sold < 1e-9 ? 'ок' : 'нарушение (нет в списке филиала)';
    },

    discountPercent: function(type){
      var text = type === null || type === undefined ? '' : String(type).trim();
      var match = /(\d{1,2})\s*%/.exec(text);
      return match ? +match[1] / 100 : null;
    },

    estimatedDiscountAmount: function(quantity, pieceFactor, price, type){
      var percent = this.discountPercent(type);
      return percent ? quantity * pieceFactor * price * percent : null;
    }
  };
  root.Calculations = calculations;
  if (typeof module !== 'undefined' && module.exports) module.exports = calculations;
})(typeof window !== 'undefined' ? window : this);
