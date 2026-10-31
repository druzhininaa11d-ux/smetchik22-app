/* ============================================================
   BuildMate Pro — патч экспорта Excel / PDF / JSON
   Версия: 1.9 (2026-10)
   Подключать ПОСЛЕ основного скрипта index.html:
     <script src="patch.js"></script>
   Переопределяет:
     - saveBlob          (универсальное скачивание файлов)
     - buildExcelBlob    (Excel с автофильтром и шириной колонок)
     - buildPdfBlob      (PDF с кириллицей через Roboto)
     - ensureXlsx        (загрузка XLSX с fallback CDN)
     - ensurePdfMake     (загрузка pdfMake + VFS с fallback CDN)
     - exportExcel       (смета в Excel)
     - exportCatalogExcel(прайс каталога в Excel)
     - exportBackup      (резервная копия JSON)
     - downloadExcelTemplate (шаблон импорта)
     - buildEstimateTableBody / printDoc (PDF-документы)
   ============================================================ */
(function () {
  'use strict';

  // ---------- Проверка, что основной скрипт уже загружен ----------
  // Все функции основного скрипта объявлены в глобальной области
  // (var/function hoisting в <script>), поэтому доступны через window.
  if (typeof window.saveBlob !== 'function' ||
      typeof window.buildExcelBlob !== 'function' ||
      typeof window.printDoc !== 'function') {
    console.warn('[patch.js] Основной скрипт не найден. Подключите patch.js ПОСЛЕ основного <script>.');
    return;
  }

  // ---------- Константы CDN ----------
  var CDN = {
    XLSX:        'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    XLSX_ALT:    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    PDFMAKE:     'https://cdn.jsdelivr.net/npm/pdfmake@0.2.7/build/pdfmake.min.js',
    PDFMAKE_ALT: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js',
    VFS:         'https://cdn.jsdelivr.net/npm/pdfmake@0.2.7/build/vfs_fonts.js',
    VFS_ALT:     'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js'
  };

  // ---------- Утилита: загрузка скрипта с таймаутом ----------
  function loadScriptOnce(src, timeoutMs) {
    timeoutMs = timeoutMs || 9000;
    return new Promise(function (resolve, reject) {
      if (document.querySelector('script[src="' + src + '"]')) return resolve();
      var s = document.createElement('script');
      s.src = src;
      s.async = true;
      var timer = setTimeout(function () {
        s.onload = s.onerror = null;
        reject(new Error('Таймаут загрузки: ' + src));
      }, timeoutMs);
      s.onload  = function () { clearTimeout(timer); resolve(); };
      s.onerror = function () { clearTimeout(timer); reject(new Error('Не загрузилось: ' + src)); };
      document.head.appendChild(s);
    });
  }

  // ---------- Загрузчики библиотек с fallback ----------
  function ensureXlsx() {
    if (window.XLSX) return Promise.resolve();
    return loadScriptOnce(CDN.XLSX)
      .catch(function () { return loadScriptOnce(CDN.XLSX_ALT); })
      .then(function () {
        if (!window.XLSX) throw new Error('Библиотека Excel не загрузилась. Проверьте интернет.');
      });
  }

  function ensurePdfMake() {
    if (window.pdfMake && window.pdfMake.vfs && Object.keys(window.pdfMake.vfs).length) {
      // Уже загружено — просто настраиваем шрифты
      window.pdfMake.fonts = {
        Roboto: {
          normal: 'Roboto-Regular.ttf',
          bold: 'Roboto-Medium.ttf',
          italics: 'Roboto-Italic.ttf',
          bolditalics: 'Roboto-MediumItalic.ttf'
        }
      };
      return Promise.resolve();
    }
    var p = Promise.resolve();
    if (!window.pdfMake) {
      p = p.then(function () {
        return loadScriptOnce(CDN.PDFMAKE).catch(function () {
          return loadScriptOnce(CDN.PDFMAKE_ALT);
        });
      });
    }
    return p
      .then(function () {
        if (!window.pdfMake) throw new Error('pdfMake не загрузился');
        if (!window.pdfMake.vfs || !Object.keys(window.pdfMake.vfs).length) {
          return loadScriptOnce(CDN.VFS).catch(function () {
            return loadScriptOnce(CDN.VFS_ALT);
          });
        }
      })
      .then(function () {
        if (!window.pdfMake.vfs || !Object.keys(window.pdfMake.vfs).length) {
          throw new Error('Шрифты PDF не загрузились');
        }
        window.pdfMake.fonts = {
          Roboto: {
            normal: 'Roboto-Regular.ttf',
            bold: 'Roboto-Medium.ttf',
            italics: 'Roboto-Italic.ttf',
            bolditalics: 'Roboto-MediumItalic.ttf'
          }
        };
      });
  }

  // ---------- Blob → Data URL ----------
  function blobToDataURL(blob) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload  = function () { resolve(r.result); };
      r.onerror = reject;
      r.readAsDataURL(blob);
    });
  }

  // ---------- Универсальное сохранение файла ----------
  // Каскад: Web Share → Telegram downloadFile → a.download → открытие Data URL
  function saveBlob(blob, filename) {
    // 1) Web Share API (iOS / Android / Telegram WebView)
    return Promise.resolve()
      .then(function () {
        if (navigator.canShare && navigator.share) {
          var file = new File([blob], filename, { type: blob.type || 'application/octet-stream' });
          if (navigator.canShare({ files: [file] })) {
            return navigator.share({ files: [file], title: filename })
              .then(function () {
                if (window.toast) window.toast('✅ Файл передан в «Поделиться»');
                return true;
              });
          }
        }
      })
      .catch(function (e) {
        if (e && e.name === 'AbortError') {
          if (window.toast) window.toast('Отменено', '#64748b');
          return true; // не пробуем другие способы — пользователь сам отменил
        }
        console.warn('[patch.js] Web Share failed:', e);
      })
      // 2) Telegram downloadFile
      .then(function (result) {
        if (result === true) return true;
        var isTelegram = !!(window.Telegram && window.Telegram.WebApp && window.Telegram.WebApp.initData);
        if (isTelegram && typeof window.Telegram.WebApp.downloadFile === 'function') {
          return blobToDataURL(blob).then(function (dataUrl) {
            window.Telegram.WebApp.downloadFile({ url: dataUrl, file_name: filename });
            if (window.toast) window.toast('✅ Файл передан на скачивание');
            return true;
          }).catch(function (e) {
            console.warn('[patch.js] Telegram.downloadFile failed:', e);
            return false;
          });
        }
        return false;
      })
      // 3) a.download
      .then(function (result) {
        if (result === true) return true;
        try {
          var url = URL.createObjectURL(blob);
          var a = document.createElement('a');
          a.href = url;
          a.download = filename;
          a.rel = 'noopener';
          a.style.display = 'none';
          document.body.appendChild(a);
          a.click();
          setTimeout(function () {
            if (a.parentNode) a.parentNode.removeChild(a);
            try { URL.revokeObjectURL(url); } catch (e) {}
          }, 2000);
          if (window.toast) window.toast('✅ Файл скачан');
          return true;
        } catch (e) {
          console.warn('[patch.js] a.download failed:', e);
          return false;
        }
      })
      // 4) Открытие Data URL в новом окне
      .then(function (result) {
        if (result === true) return true;
        return blobToDataURL(blob).then(function (dataUrl) {
          var isTelegram = !!(window.Telegram && window.Telegram.WebApp && window.Telegram.WebApp.initData);
          if (isTelegram && window.Telegram.WebApp.openLink) {
            window.Telegram.WebApp.openLink(dataUrl);
          } else {
            var w = window.open(dataUrl, '_blank');
            if (!w) throw new Error('popup blocked');
          }
          if (window.toast) window.toast('📥 Файл открыт. Сохраните через меню браузера.', '#64748b');
          return true;
        });
      })
      .catch(function (e) {
        console.warn('[patch.js] Open fallback failed:', e);
        alert('Не удалось сохранить файл. Попробуйте открыть приложение в браузере.');
        return false;
      });
  }

  // ---------- Excel: сборка Blob с автофильтром и шириной ----------
  function buildExcelBlob(rows, sheetName) {
    return ensureXlsx().then(function () {
      var ws = XLSX.utils.aoa_to_sheet(rows);

      // Ширина колонок с учётом кириллицы
      var colWidths = [];
      rows.forEach(function (row) {
        (row || []).forEach(function (cell, idx) {
          var str = String(cell == null ? '' : cell);
          var weight = /[А-Яа-яЁё]/.test(str) ? 1.3 : 1;
          var w = Math.min(60, Math.ceil(str.length * weight) + 2);
          colWidths[idx] = Math.max(colWidths[idx] || 8, w);
        });
      });
      ws['!cols'] = colWidths.map(function (w) { return { wch: w }; });

      // Автофильтр по шапке
      if (rows.length > 1 && ws['!ref']) {
        var range = XLSX.utils.decode_range(ws['!ref']);
        ws['!autofilter'] = {
          ref: XLSX.utils.encode_range({
            s: { r: range.s.r, c: range.s.c },
            e: { r: range.s.r, c: range.e.c }
          })
        };
      }

      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, (sheetName || 'Смета').slice(0, 31));

      return new Promise(function (resolve, reject) {
        try {
          XLSX.write(wb, { bookType: 'xlsx', type: 'array' }, function (arr) {
            resolve(new Blob([arr], {
              type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }));
          });
        } catch (e) { reject(e); }
      });
    });
  }

  // ---------- PDF: сборка Blob с кириллицей ----------
  function buildPdfBlob(docDefinition) {
    return ensurePdfMake().then(function () {
      // Гарантируем Roboto и базовый стиль
      docDefinition.defaultStyle = Object.assign(
        { font: 'Roboto', fontSize: 9, lineHeight: 1.3 },
        docDefinition.defaultStyle || {}
      );
      return new Promise(function (resolve, reject) {
        try {
          pdfMake.createPdf(docDefinition).getBlob(
            function (blob) {
              if (!blob || blob.size < 100) return reject(new Error('Пустой PDF'));
              resolve(new Blob([blob], { type: 'application/pdf' }));
            },
            function (err) { reject(err || new Error('Ошибка генерации PDF')); }
          );
        } catch (e) { reject(e); }
      });
    });
  }

  // ---------- Excel: смета ----------
  function exportExcel() {
    var o = window.obj ? window.obj() : null;
    if (!o) { alert('Выберите объект'); return Promise.resolve(); }

    if (window.toast) window.toast('⏳ Формирую Excel...', '#64748b');
    var items = window.getEstimateItems ? window.getEstimateItems(o) : o.items;
    var t = window.totals ? window.totals(o, items) : { work: 0, mat: 0, markup: 0, total: 0 };
    var today = window.today ? window.today() : new Date().toISOString().slice(0, 10);
    var num = window.num || function (n) { return Math.max(0, Number(n) || 0); };

    var rows = [
      ['Объект', o.name], ['Тип', o.kind || ''], ['Адрес', o.address || ''],
      ['Заказчик', o.client || ''], ['Дата', today], [],
      ['№', 'Тип', 'Помещение', 'Этап', 'Раздел', 'Наименование',
       'Кол-во', 'Ед.', 'Цена', 'Сумма', 'Себестоимость']
    ];
    items.forEach(function (i, idx) {
      var room = i.roomId ? (o.rooms || []).find(function (r) { return r.id === i.roomId; }) : null;
      rows.push([
        idx + 1,
        i.type === 'material' ? 'Материал' : 'Работа',
        room ? room.name : (i.roomName || 'Общее'),
        i.stage || '', i.group || '', i.name || '',
        Number(num(i.qty)), i.unit || '',
        Number(num(i.price)), Number(num(i.qty) * num(i.price)),
        Number(num(i.cost))
      ]);
    });
    rows.push([]);
    rows.push(['', '', '', '', '', 'Работы', '', '', '', Number(t.work)]);
    rows.push(['', '', '', '', '', 'Материалы', '', '', '', Number(t.mat)]);
    rows.push(['', '', '', '', '', 'Наценка ' + (window.db ? num(window.db.markup) : 0) + '%', '', '', '', Number(t.markup)]);
    rows.push(['', '', '', '', '', 'ИТОГО', '', '', '', Number(t.total)]);

    return buildExcelBlob(rows, 'Смета')
      .then(function (blob) {
        return saveBlob(blob, 'buildmate-smeta-' + (o.contractNo || o.id) + '-' + today + '.xlsx');
      })
      .catch(function (e) {
        console.error(e);
        alert('Ошибка при создании Excel: ' + (e.message || e));
      });
  }

  // ---------- Excel: прайс каталога ----------
  function exportCatalogExcel() {
    var CATALOG = window.CATALOG || [];
    var today = window.today ? window.today() : new Date().toISOString().slice(0, 10);
    var num = window.num || function (n) { return Math.max(0, Number(n) || 0); };

    if (window.toast) window.toast('⏳ Формирую Excel...', '#64748b');

    var rows = [['Тип', 'Этап', 'Раздел', 'Наименование', 'Единица', 'Цена', 'Помещения']];
    CATALOG.forEach(function (p) {
      rows.push([
        p.type === 'material' ? 'Материал' : 'Работа',
        p.stage || '', p.group || '', p.name || '', p.unit || '',
        Number(num(p.price)),
        (p.rooms || []).join(',')
      ]);
    });

    return buildExcelBlob(rows, 'Расценки')
      .then(function (blob) {
        return saveBlob(blob, 'buildmate-catalog-' + today + '.xlsx');
      })
      .catch(function (e) {
        alert('Ошибка: ' + (e.message || e));
      });
  }

  // ---------- JSON: резервная копия ----------
  function exportBackup() {
    if (window.saveTimeout) { clearTimeout(window.saveTimeout); window.save && window.save(); window.saveTimeout = null; }
    if (window.db && window.CATALOG) window.db.catalog = window.CATALOG;
    var today = window.today ? window.today() : new Date().toISOString().slice(0, 10);
    try {
      var json = JSON.stringify(window.db, null, 2);
      var blob = new Blob([json], { type: 'application/json;charset=utf-8' });
      return saveBlob(blob, 'buildmate-backup-' + today + '.json');
    } catch (e) {
      alert('Ошибка: ' + (e.message || e));
      return Promise.resolve();
    }
  }

  // ---------- Excel: шаблон импорта ----------
  function downloadExcelTemplate() {
    return ensureXlsx().then(function () {
      var data = [
        ['Наименование', 'Тип', 'Раздел', 'Единица', 'Цена', 'Этап', 'Помещения'],
        ['Штукатурка стен по маякам', 'Работа', 'Штукатурка', 'м²', 650, 'Черновые', 'Спальня,Гостиная,Кухня'],
        ['Шпаклёвка стен под обои', 'Работа', 'Шпаклёвка', 'м²', 280, 'Черновые', 'Спальня,Гостиная'],
        ['Плитка керамическая', 'Материал', 'Плитка', 'м²', 1700, 'Чистовые', 'Ванная,Санузел'],
        ['Унитаз с инсталляцией', 'Материал', 'Сантехника', 'комплект', 20000, 'Чистовые', 'Санузел'],
        ['', '', '', '', '', '', '']
      ];
      return buildExcelBlob(data, 'Расценки');
    }).then(function (blob) {
      return saveBlob(blob, 'buildmate-pro-template.xlsx');
    }).catch(function (e) {
      alert('Не удалось создать шаблон: ' + (e.message || e));
    });
  }

  // ---------- PDF: таблица сметы ----------
  function buildEstimateTableBody(items, o) {
    var num = window.num || function (n) { return Math.max(0, Number(n) || 0); };
    var fmt = window.fmt || function (n) { return String(n) + ' ₽'; };

    var body = [[
      { text: '№', bold: true, style: 'th', alignment: 'center' },
      { text: 'Наименование', bold: true, style: 'th' },
      { text: 'Ед.', bold: true, style: 'th', alignment: 'center' },
      { text: 'Кол-во', bold: true, style: 'th', alignment: 'right' },
      { text: 'Цена', bold: true, style: 'th', alignment: 'right' },
      { text: 'Сумма', bold: true, style: 'th', alignment: 'right' }
    ]];
    items.forEach(function (i, n) {
      var room = i.roomId ? (o.rooms || []).find(function (r) { return r.id === i.roomId; }) : null;
      var label = room ? room.name : (i.roomName || '');
      var name = label ? i.name + ' (' + label + ')' : i.name;
      body.push([
        { text: String(n + 1), alignment: 'center' },
        { text: name },
        { text: i.unit || '', alignment: 'center' },
        { text: String(num(i.qty)), alignment: 'right' },
        { text: fmt(i.price), alignment: 'right' },
        { text: fmt(num(i.qty) * num(i.price)), alignment: 'right' }
      ]);
    });
    return body;
  }

  // ---------- PDF: документ ----------
  function printDoc(kind) {
    var o = window.obj ? window.obj() : null;
    if (!o) { alert('Выберите объект'); return Promise.resolve(); }

    var db = window.db || {};
    var today = window.today ? window.today() : new Date().toISOString().slice(0, 10);
    var fmt = window.fmt || function (n) { return String(n) + ' ₽'; };
    var totals = window.totals || function () { return { work: 0, mat: 0, markup: 0, total: 0 }; };
    var companyRequisites = window.companyRequisites || function () { return ''; };

    if (window.toast) window.toast('⏳ Формирую PDF...', '#64748b');

    var t = totals(o);
    var items = o.items;
    var reqs = companyRequisites();
    var companyName = db.company && db.company.trim() ? db.company : '________________________';
    var city = db.region || 'Новосибирск';
    var title = kind === 'contract' ? 'Договор оказания услуг'
              : kind === 'act'      ? 'Акт выполненных работ'
              : kind === 'invoice'  ? 'Счёт на оплату'
              : 'Смета на ремонт';

    var tableBody = buildEstimateTableBody(items, o);
    var content = [];

    if (kind === 'contract') {
      content.push({ text: title + ' № ' + (o.contractNo || ''), style: 'h1' });
      content.push({ text: 'г. ' + city + ' «' + (o.contractDate || today) + '»', margin: [0, 8, 0, 12] });
      content.push({ text: 'Исполнитель: ' + (reqs || '________________________') + ', Заказчик: ' + (o.client || '________________________') + ' заключили настоящий договор.', margin: [0, 0, 0, 10] });
      content.push({ text: '1. Предмет договора', style: 'h3' });
      content.push({ text: 'Исполнитель выполняет ремонтные работы на объекте: ' + o.name + ' (' + (o.kind || '—') + ') по адресу: ' + (o.address || '________________________') + '.' });
      content.push({ text: '2. Стоимость и порядок оплаты', style: 'h3' });
      content.push({ text: 'Стоимость по смете: ' + fmt(t.total) + '. ' + (o.contractTerms || '') });
      content.push({ text: '3. Приёмка', style: 'h3' });
      content.push({ text: 'Порядок и сроки согласовываются дополнительно.' });
      content.push({ text: '4. Подписи сторон', style: 'h3' });
      content.push({ text: 'Исполнитель: ' + companyName + (db.director ? ' (' + db.director + ')' : '') + ' ___________', margin: [0, 8, 0, 0] });
      content.push({ text: 'Заказчик: ' + (o.client || '________________') + ' ___________', margin: [0, 8, 0, 0] });
      content.push({ text: 'Приложение № 1. Смета', style: 'h3', pageBreak: 'before' });
    } else if (kind === 'act') {
      content.push({ text: title, style: 'h1' });
      content.push({ text: 'Объект: ' + o.name + ' (' + (o.kind || '—') + '). Адрес: ' + (o.address || '—'), margin: [0, 8, 0, 6] });
      content.push({ text: 'Исполнитель: ' + (reqs || '________________________') });
      content.push({ text: 'Приняты работы на сумму ' + fmt(t.total) + '.', margin: [0, 8, 0, 12] });
    } else if (kind === 'invoice') {
      content.push({ text: title + ' № ' + (o.contractNo || '') + ' от ' + (o.contractDate || today), style: 'h1' });
      content.push({ text: 'Исполнитель: ' + (reqs || '________________________'), margin: [0, 8, 0, 4] });
      content.push({ text: 'Заказчик: ' + (o.client || '—'), margin: [0, 0, 0, 12] });
    } else {
      content.push({ text: title, style: 'h1' });
      content.push({ text: 'Компания: ' + (reqs || '________________________') + ' · Объект: ' + o.name + ' (' + (o.kind || '—') + ') · Адрес: ' + (o.address || '—'), margin: [0, 8, 0, 4], fontSize: 10 });
      content.push({ text: 'Заказчик: ' + (o.client || '—') + ' · Дата: ' + today, margin: [0, 0, 0, 12], fontSize: 10 });
    }

    content.push({
      table: { headerRows: 1, widths: [18, '*', 28, 40, 60, 70], dontBreakRows: true, body: tableBody },
      layout: {
        hLineWidth: function () { return 0.5; },
        vLineWidth: function () { return 0.5; },
        hLineColor: function () { return '#cccccc'; },
        vLineColor: function () { return '#cccccc'; },
        paddingLeft: function () { return 4; },
        paddingRight: function () { return 4; },
        paddingTop: function () { return 3; },
        paddingBottom: function () { return 3; }
      },
      margin: [0, 10, 0, 0]
    });
    content.push({ text: 'Работы: ' + fmt(t.work), alignment: 'right', margin: [0, 12, 0, 2], fontSize: 11 });
    content.push({ text: 'Материалы: ' + fmt(t.mat), alignment: 'right', margin: [0, 0, 0, 2], fontSize: 11 });
    content.push({ text: 'Наценка: ' + fmt(t.markup), alignment: 'right', margin: [0, 0, 0, 2], fontSize: 11 });
    content.push({ text: 'ИТОГО: ' + fmt(t.total), alignment: 'right', bold: true, fontSize: 14, margin: [0, 6, 0, 0] });

    if (kind === 'act') {
      content.push({ text: 'Замечания: _______________', margin: [0, 20, 0, 0] });
      content.push({ text: 'Исполнитель ' + (db.director ? '(' + db.director + ')' : '') + ' __________   Заказчик __________', margin: [0, 20, 0, 0] });
    } else if (kind === 'invoice') {
      content.push({ text: 'К оплате: ' + fmt(t.total), bold: true, margin: [0, 16, 0, 0] });
      content.push({ text: 'Руководитель ' + (db.director ? '(' + db.director + ')' : '') + ' ___________', margin: [0, 20, 0, 0] });
    }
    content.push({ text: 'Сформировано в BuildMate Pro · ' + today, fontSize: 8, color: '#888888', margin: [0, 20, 0, 0], alignment: 'right' });

    var docDef = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      footer: function (cur, total) { return { text: cur + ' / ' + total, alignment: 'center', fontSize: 8, color: '#888' }; },
      content: content,
      styles: {
        h1: { fontSize: 16, bold: true, alignment: 'center', margin: [0, 0, 0, 8] },
        h3: { fontSize: 12, bold: true, margin: [0, 12, 0, 4] },
        th: { fillColor: '#f0f0f0', fontSize: 9, bold: true }
      },
      defaultStyle: { font: 'Roboto', fontSize: 9, lineHeight: 1.3 }
    };

    return buildPdfBlob(docDef)
      .then(function (blob) {
        return saveBlob(blob, 'buildmate-' + kind + '-' + (o.contractNo || o.id) + '-' + today + '.pdf');
      })
      .catch(function (e) {
        console.error(e);
        alert('Ошибка при создании PDF: ' + (e.message || e));
      });
  }

  // ---------- Замена импорта Excel (совместимость) ----------
  // В основном файле функция ensureXlsx() используется при импорте.
  // Перенаправляем её на нашу реализацию с fallback CDN.
  window.ensureXlsxImport = ensureXlsx;

  // ---------- Подмена глобальных функций ----------
  window.saveBlob            = saveBlob;
  window.buildExcelBlob      = buildExcelBlob;
  window.buildPdfBlob        = buildPdfBlob;
  window.ensureXlsx          = ensureXlsx;
  window.ensurePdfMake       = ensurePdfMake;
  window.exportExcel         = exportExcel;
  window.exportCatalogExcel  = exportCatalogExcel;
  window.exportBackup        = exportBackup;
  window.downloadExcelTemplate = downloadExcelTemplate;
  window.buildEstimateTableBody = buildEstimateTableBody;
  window.printDoc            = printDoc;

  console.log('✅ [patch.js] Патч экспорта Excel/PDF/JSON загружен (v1.9)');
})();
