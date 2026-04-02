/**
 * 角色初印象表格 — 左预览 / 右编辑、裁剪、双列富文本、导出
 */

(function () {
  'use strict';

  var STORAGE_KEY = 'characterImpressionTable_v2';
  var CROP_MAX_EDGE = 1200;

  /**
   * 社交平台常见导出尺寸（宽×高，像素）。mode: contain=整图缩放留白；fitWidth=固定宽、高随内容。
   */
  var EXPORT_PRESETS = {
    xhs_3_4: { label: '小红书 笔记 3:4', w: 1080, h: 1440, mode: 'contain' },
    xhs_4_5: { label: '小红书 竖图 4:5', w: 1080, h: 1350, mode: 'contain' },
    xhs_1_1: { label: '小红书 / 微博 方形 1:1', w: 1080, h: 1080, mode: 'contain' },
    weibo_long: { label: '微博 长图（宽 1080）', w: 1080, h: null, mode: 'fitWidth' },
    story_9_16: { label: '竖屏动态 9:16', w: 1080, h: 1920, mode: 'contain' },
  };

  /** @type {{ workTitle: string, filledBy: string, madeBy: string, bgBaseTone: string, bgImageSrc: string, bgOpacity: number, imageColWidthPct: number, exportPreset: string }} */
  var meta = {};

  /** @type {{ imgSrc: string, richFirst: string, richCurrent: string, rowMinHeight: number }[]} */
  var tableData = [];

  var currentRichCell = null;
  var pendingUploadRowIndex = -1;
  var pendingImageActionRowIndex = -1;
  var cropRowIndex = -1;
  /** @type {{ destroy: function(): void, getCroppedCanvas: function(Object): HTMLCanvasElement }|null} */
  var cropperInstance = null;
  var savedSelectionRange = null;
  /** 裁剪弹窗使用临时图（先裁后写入） */
  var cropForceSrc = null;
  /** 列宽拖动 */
  var colDrag = null;
  /** 行高拖动 */
  var rowDrag = null;

  function getDefaultMeta() {
    return {
      workTitle: '',
      filledBy: '',
      madeBy: '',
      bgBaseTone: 'white',
      bgImageSrc: '',
      bgOpacity: 40,
      imageColWidthPct: 38,
      exportPreset: 'xhs_3_4',
    };
  }

  function defaultRows() {
    return [
      normalizeRow({
        imgSrc: '',
        richFirst: '<p><strong>示例角色</strong> 这是初印象描述...</p>',
        richCurrent: '<p></p>',
        rowMinHeight: 0,
      }),
      normalizeRow({
        imgSrc: '',
        richFirst: '<p>新角色</p>',
        richCurrent: '<p></p>',
        rowMinHeight: 0,
      }),
    ];
  }

  var tableBody = document.getElementById('tableBody');
  var mainTable = document.getElementById('mainTable');
  var colImageWidthCol = document.getElementById('colImageWidth');
  var colResizeHandle = document.getElementById('colResizeHandle');
  var captureArea = document.getElementById('captureArea');
  var previewBgLayer = document.getElementById('previewBgLayer');
  var workTitle = document.getElementById('workTitle');
  var filledBy = document.getElementById('filledBy');
  var madeBy = document.getElementById('madeBy');
  var bgBaseTone = document.getElementById('bgBaseTone');
  var bgOpacity = document.getElementById('bgOpacity');
  var bgOpacityVal = document.getElementById('bgOpacityVal');
  var btnBgImage = document.getElementById('btnBgImage');
  var btnClearBg = document.getElementById('btnClearBg');
  var btnAddRow = document.getElementById('btnAddRow');
  var btnRemoveRow = document.getElementById('btnRemoveRow');
  var btnExport = document.getElementById('btnExport');
  var exportFormat = document.getElementById('exportFormat');
  var exportPreset = document.getElementById('exportPreset');
  var previewPresetBadge = document.getElementById('previewPresetBadge');
  var previewDeviceScreen = document.getElementById('previewDeviceScreen');
  var hiddenFileInput = document.getElementById('hiddenFileInput');
  var hiddenBgInput = document.getElementById('hiddenBgInput');
  var btnBold = document.getElementById('btnBold');
  var btnItalic = document.getElementById('btnItalic');
  var btnUnderline = document.getElementById('btnUnderline');
  var btnFontInc = document.getElementById('btnFontInc');
  var btnFontDec = document.getElementById('btnFontDec');
  var colorPicker = document.getElementById('colorPicker');
  var imageActionModal = document.getElementById('imageActionModal');
  var btnImageReplace = document.getElementById('btnImageReplace');
  var btnImageCrop = document.getElementById('btnImageCrop');
  var btnImageActionCancel = document.getElementById('btnImageActionCancel');
  var cropModal = document.getElementById('cropModal');
  var cropTarget = document.getElementById('cropTarget');
  var btnCropCancel = document.getElementById('btnCropCancel');
  var btnCropConfirm = document.getElementById('btnCropConfirm');

  // 添加移动端导出图片模态框相关元素
  var exportedImageModal = document.getElementById('exportedImageModal');
  var exportedImageDisplay = document.getElementById('exportedImageDisplay');
  var btnCloseExportModal = document.getElementById('btnCloseExportModal');

  var PL_FIRST = '双击编辑初印象…';

  function saveToLocalStorage() {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({ meta: meta, rows: tableData })
      );
    } catch (e) {
      console.warn('localStorage 保存失败:', e);
    }
  }

  function normalizeRow(row) {
    var first =
      row.richFirst != null
        ? row.richFirst
        : row.richText != null
          ? row.richText
          : '<p></p>';
    var current = row.richCurrent != null ? row.richCurrent : '<p></p>';
    var rh = row.rowMinHeight;
    if (typeof rh !== 'number' || isNaN(rh) || rh < 0) rh = 0;
    if (rh > 1600) rh = 1600;
    return {
      imgSrc: typeof row.imgSrc === 'string' ? row.imgSrc : '',
      richFirst: typeof first === 'string' ? first : '<p></p>',
      richCurrent: typeof current === 'string' ? current : '<p></p>',
      rowMinHeight: Math.round(rh),
    };
  }

  function normalizeMeta(m) {
    var d = getDefaultMeta();
    if (!m || typeof m !== 'object') return d;
    return {
      workTitle: typeof m.workTitle === 'string' ? m.workTitle : d.workTitle,
      filledBy: typeof m.filledBy === 'string' ? m.filledBy : d.filledBy,
      madeBy: typeof m.madeBy === 'string' ? m.madeBy : d.madeBy,
      bgBaseTone: m.bgBaseTone === 'black' ? 'black' : 'white',
      bgImageSrc: typeof m.bgImageSrc === 'string' ? m.bgImageSrc : '',
      bgOpacity:
        typeof m.bgOpacity === 'number' && m.bgOpacity >= 0 && m.bgOpacity <= 100
          ? Math.round(m.bgOpacity)
          : d.bgOpacity,
      imageColWidthPct:
        typeof m.imageColWidthPct === 'number' && m.imageColWidthPct >= 22 && m.imageColWidthPct <= 62
          ? Math.round(m.imageColWidthPct * 10) / 10
          : d.imageColWidthPct,
      exportPreset:
        typeof m.exportPreset === 'string' && EXPORT_PRESETS[m.exportPreset]
          ? m.exportPreset
          : d.exportPreset,
    };
  }

  function getActiveExportPreset() {
    var id = meta.exportPreset;
    if (!EXPORT_PRESETS[id]) id = 'xhs_3_4';
    var p = EXPORT_PRESETS[id];
    return { id: id, label: p.label, w: p.w, h: p.h, mode: p.mode };
  }

  function getExportDimensions(presetId) {
    var preset = EXPORT_PRESETS[presetId];
    if (!preset) {
      preset = EXPORT_PRESETS['xhs_3_4']; // 默认值
    }
    
    if (preset.mode === 'fitWidth' || preset.h == null) {
      // 对于微博长图模式，高度应根据内容确定
      return { width: preset.w, height: null };
    }
    
    // 对于其他模式，返回固定宽高
    return { width: preset.w, height: preset.h };
  }

  /** 将 html2canvas 结果缩放为预设像素（contain 留白 / fitWidth 仅限宽） */
  function resizeCanvasToExport(srcCanvas, preset, bgHex) {
    var sw = srcCanvas.width;
    var sh = srcCanvas.height;
    if (sw <= 0 || sh <= 0) return srcCanvas;

    if (preset.mode === 'fitWidth' || preset.h == null) {
      var tw = preset.w;
      var th = Math.max(1, Math.round(sh * (tw / sw)));
      var outW = document.createElement('canvas');
      outW.width = tw;
      outW.height = th;
      var ctxW = outW.getContext('2d');
      if (!ctxW) return srcCanvas;
      ctxW.fillStyle = bgHex;
      ctxW.fillRect(0, 0, tw, th);
      ctxW.drawImage(srcCanvas, 0, 0, sw, sh, 0, 0, tw, th);
      return outW;
    }

    var tw = preset.w;
    var th = preset.h;
    var scale = Math.min(tw / sw, th / sh);
    var dw = Math.max(1, Math.round(sw * scale));
    var dh = Math.max(1, Math.round(sh * scale));
    var ox = Math.floor((tw - dw) / 2);
    var oy = Math.floor((th - dh) / 2);
    var out = document.createElement('canvas');
    out.width = tw;
    out.height = th;
    var ctx = out.getContext('2d');
    if (!ctx) return srcCanvas;
    ctx.fillStyle = bgHex;
    ctx.fillRect(0, 0, tw, th);
    ctx.drawImage(srcCanvas, 0, 0, sw, sh, ox, oy, dw, dh);
    return out;
  }

  function applyExportPresetFrame() {
    var p = getActiveExportPreset();
    if (previewPresetBadge) {
      previewPresetBadge.textContent =
        p.mode === 'fitWidth'
          ? p.label + ' · 宽 ' + p.w + 'px（高随内容）'
          : p.label + ' · ' + p.w + '×' + p.h;
    }
    if (previewDeviceScreen) {
      previewDeviceScreen.classList.toggle('preview-screen--fluid', p.mode === 'fitWidth');
      if (p.mode === 'fitWidth') {
        previewDeviceScreen.style.aspectRatio = '';
      } else {
        previewDeviceScreen.style.aspectRatio = String(p.w) + ' / ' + String(p.h);
      }
    }
  }

  function loadFromLocalStorage() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return { meta: getDefaultMeta(), rows: parsed.map(normalizeRow) };
      }
      if (parsed && Array.isArray(parsed.rows)) {
        return {
          meta: normalizeMeta(parsed.meta),
          rows: parsed.rows.length ? parsed.rows.map(normalizeRow) : defaultRows(),
        };
      }
    } catch (e) {
      console.warn('localStorage 读取失败:', e);
    }
    return null;
  }

  /** 兼容旧版仅保存行数组的 v1 数据 */
  function loadLegacyV1() {
    try {
      var raw = localStorage.getItem('characterImpressionTable_v1');
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      if (!Array.isArray(parsed) || !parsed.length) return null;
      return { meta: getDefaultMeta(), rows: parsed.map(normalizeRow) };
    } catch (e) {
      return null;
    }
  }

  function pushMetaToInputs() {
    if (workTitle) workTitle.value = meta.workTitle;
    if (filledBy) filledBy.value = meta.filledBy;
    if (madeBy) madeBy.value = meta.madeBy;
    if (bgBaseTone) bgBaseTone.value = meta.bgBaseTone;
    if (bgOpacity) bgOpacity.value = String(meta.bgOpacity);
    if (bgOpacityVal) bgOpacityVal.textContent = meta.bgOpacity + '%';
    if (exportPreset) exportPreset.value = meta.exportPreset;
    applyImageColPct();
  }

  function applyImageColPct() {
    var pct = meta.imageColWidthPct;
    var rest = Math.max(38, 100 - pct);
    if (colImageWidthCol) colImageWidthCol.style.width = pct + '%';
    var textCol = mainTable && mainTable.querySelector('.col-w-text');
    if (textCol) textCol.style.width = rest + '%';
  }

  function applyPreviewStyles() {
    if (!captureArea) return;
    captureArea.classList.toggle('sheet-tone-black', meta.bgBaseTone === 'black');
    if (previewBgLayer) {
      if (meta.bgImageSrc) {
        previewBgLayer.style.backgroundImage =
          'url(' + JSON.stringify(meta.bgImageSrc) + ')';
        previewBgLayer.style.display = 'block';
      } else {
        previewBgLayer.style.backgroundImage = 'none';
        previewBgLayer.style.display = 'none';
      }
      previewBgLayer.style.opacity = String(meta.bgOpacity / 100);
    }
  }

  function updateRemoveRowButton() {
    if (btnRemoveRow) {
      btnRemoveRow.disabled = tableData.length <= 1;
    }
  }

  function isRichCellVisuallyEmpty(el) {
    if (!el) return true;
    var t = el.innerText || '';
    return t.replace(/\u200b/g, '').trim().length === 0;
  }

  function updateRichEmptyClass(el) {
    if (!el) return;
    el.classList.toggle('is-empty', isRichCellVisuallyEmpty(el));
  }

  function syncRichTextFromCell(rowIndex, el) {
    if (!el || rowIndex < 0 || rowIndex >= tableData.length) return;
    var field = el.getAttribute('data-rich-field');
    if (field === 'current') {
      tableData[rowIndex].richCurrent = el.innerHTML;
    } else {
      tableData[rowIndex].richFirst = el.innerHTML;
    }
    saveToLocalStorage();
  }

  function createPlaceholderSvg() {
    var ns = 'http://www.w3.org/2000/svg';
    var svg = document.createElementNS(ns, 'svg');
    svg.setAttribute('width', '100');
    svg.setAttribute('height', '88');
    svg.setAttribute('viewBox', '0 0 100 88');
    svg.setAttribute('aria-hidden', 'true');
    var rect = document.createElementNS(ns, 'rect');
    rect.setAttribute('x', '8');
    rect.setAttribute('y', '6');
    rect.setAttribute('width', '84');
    rect.setAttribute('height', '76');
    rect.setAttribute('rx', '6');
    rect.setAttribute('fill', 'currentColor');
    rect.setAttribute('opacity', '0.12');
    svg.appendChild(rect);
    return svg;
  }

  function attachRichCell(rowIndex, field, html, placeholder) {
    var rich = document.createElement('div');
    rich.className = 'rich-cell';
    rich.contentEditable = 'true';
    rich.dataset.rowIndex = String(rowIndex);
    rich.dataset.richField = field;
    rich.dataset.placeholder = placeholder;
    rich.innerHTML = html;
    updateRichEmptyClass(rich);

    rich.addEventListener('focus', function () {
      currentRichCell = rich;
    });

    rich.addEventListener('blur', function () {
      updateRichEmptyClass(rich);
      syncRichTextFromCell(rowIndex, rich);
    });

    rich.addEventListener('input', function () {
      updateRichEmptyClass(rich);
      syncRichTextFromCell(rowIndex, rich);
    });

    return rich;
  }

  function renderTable() {
    if (!tableBody) return;
    tableBody.innerHTML = '';
    applyImageColPct();

    tableData.forEach(function (row, rowIndex) {
      var tr = document.createElement('tr');
      tr.dataset.rowIndex = String(rowIndex);
      if (row.rowMinHeight > 0) {
        tr.style.minHeight = row.rowMinHeight + 'px';
      }

      var tdImg = document.createElement('td');
      tdImg.className = 'cell-image';
      if (row.rowMinHeight > 0) tdImg.style.minHeight = row.rowMinHeight + 'px';
      var inner = document.createElement('div');
      inner.className = 'cell-image-inner';

      var slot = document.createElement('div');
      slot.className = 'image-slot';

      if (row.imgSrc) {
        var img = document.createElement('img');
        img.src = row.imgSrc;
        img.alt = '角色';
        img.addEventListener('click', function () {
          openImageActionModal(rowIndex);
        });
        slot.appendChild(img);
      } else {
        var ph = document.createElement('div');
        ph.className = 'placeholder-upload';
        ph.setAttribute('role', 'button');
        ph.setAttribute('tabindex', '0');
        ph.appendChild(createPlaceholderSvg());
        var hint = document.createElement('span');
        hint.textContent = '点击上传（将先裁剪）';
        ph.appendChild(hint);
        ph.addEventListener('click', function () {
          openFilePicker(rowIndex);
        });
        ph.addEventListener('keydown', function (ev) {
          if (ev.key === 'Enter' || ev.key === ' ') {
            ev.preventDefault();
            openFilePicker(rowIndex);
          }
        });
        slot.appendChild(ph);
      }

      var controls = document.createElement('div');
      controls.className = 'image-controls hide-on-export';

      var cropBtn = document.createElement('button');
      cropBtn.type = 'button';
      cropBtn.className = 'btn-crop-img';
      cropBtn.textContent = '重新裁剪';
      cropBtn.title = '重新裁剪当前图';
      cropBtn.disabled = !row.imgSrc;
      cropBtn.addEventListener('click', function () {
        openCropModal(rowIndex);
      });

      var delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'btn-delete-img';
      delBtn.textContent = '删除图片';
      delBtn.disabled = !row.imgSrc;
      delBtn.addEventListener('click', function () {
        tableData[rowIndex].imgSrc = '';
        saveToLocalStorage();
        renderTable();
      });

      controls.appendChild(cropBtn);
      controls.appendChild(delBtn);
      inner.appendChild(slot);
      inner.appendChild(controls);
      tdImg.appendChild(inner);
      tr.appendChild(tdImg);

      var tdFirst = document.createElement('td');
      tdFirst.className = 'cell-text';
      if (row.rowMinHeight > 0) tdFirst.style.minHeight = row.rowMinHeight + 'px';
      tdFirst.appendChild(attachRichCell(rowIndex, 'first', row.richFirst, PL_FIRST));

      var rowHandle = document.createElement('div');
      rowHandle.className = 'row-resize-handle hide-on-export';
      rowHandle.dataset.rowIndex = String(rowIndex);
      rowHandle.title = '上下拖动调整本行高度';
      tdFirst.appendChild(rowHandle);

      tr.appendChild(tdFirst);

      tableBody.appendChild(tr);
    });

    updateRemoveRowButton();
  }

  function openFilePicker(rowIndex) {
    pendingUploadRowIndex = rowIndex;
    if (hiddenFileInput) {
      hiddenFileInput.value = '';
      hiddenFileInput.click();
    }
  }

  function openImageActionModal(rowIndex) {
    if (rowIndex < 0 || rowIndex >= tableData.length || !tableData[rowIndex].imgSrc) return;
    pendingImageActionRowIndex = rowIndex;
    if (imageActionModal) {
      imageActionModal.classList.add('is-open');
      imageActionModal.setAttribute('aria-hidden', 'false');
    }
  }

  function closeImageActionModal() {
    pendingImageActionRowIndex = -1;
    if (imageActionModal) {
      imageActionModal.classList.remove('is-open');
      imageActionModal.setAttribute('aria-hidden', 'true');
      // 移除焦点以避免无障碍性问题
      const activeElement = document.activeElement;
      if (activeElement && imageActionModal.contains(activeElement)) {
        activeElement.blur();
      }
    }
  }

  function destroyCropperIfAny() {
    if (cropperInstance && typeof cropperInstance.destroy === 'function') {
      cropperInstance.destroy();
    }
    cropperInstance = null;
  }

  function closeCropModal() {
    cropForceSrc = null;
    if (cropModal) {
      cropModal.classList.remove('is-open');
      cropModal.setAttribute('aria-hidden', 'true');
      // 移除焦点以避免无障碍性问题
      const activeElement = document.activeElement;
      if (activeElement && cropModal.contains(activeElement)) {
        activeElement.blur();
      }
    }
    destroyCropperIfAny();
    if (cropTarget) {
      cropTarget.removeAttribute('src');
    }
    cropRowIndex = -1;
  }

  /**
   * @param {number} rowIndex
   * @param {string} [forcedDataUrl] 新选图时传入，先裁剪再写入该行
   */
  function openCropModal(rowIndex, forcedDataUrl) {
    // 检查 Cropper 是否可用
    if (typeof Cropper === 'undefined') {
      console.error('Cropper 未定义，尝试从 window.Cropper 获取...');
      console.log('window 对象:', typeof window);
      console.log('window.Cropper:', typeof window.Cropper);
      
      // 等待一小段时间重试
      setTimeout(function() {
        if (typeof Cropper === 'undefined' && typeof window.Cropper !== 'undefined') {
          Cropper = window.Cropper;
          openCropModal(rowIndex, forcedDataUrl);
        } else {
          alert('裁剪功能不可用，请检查网络连接后重试。');
        }
      }, 500);
      return;
    }
    
    if (rowIndex < 0 || rowIndex >= tableData.length) return;
    var src =
      forcedDataUrl != null && forcedDataUrl !== ''
        ? forcedDataUrl
        : tableData[rowIndex].imgSrc;
    if (!src) return;
    cropForceSrc = forcedDataUrl != null && forcedDataUrl !== '' ? forcedDataUrl : null;
    cropRowIndex = rowIndex;
    if (!cropTarget || !cropModal) return;

    destroyCropperIfAny();
    cropModal.classList.add('is-open');
    cropModal.setAttribute('aria-hidden', 'false');

    function startCropper() {
      if (!cropModal.classList.contains('is-open')) return;
      destroyCropperIfAny();
      try {
        cropperInstance = new Cropper(cropTarget, {
          viewMode: 1,
          dragMode: 'move',
          autoCropArea: 0.85,
          restore: false,
          guides: true,
          center: true,
          highlight: false,
          cropBoxMovable: true,
          cropBoxResizable: true,
          toggleDragModeOnDblclick: false,
          background: true,
          responsive: true,
        });
      } catch (e) {
        console.error(e);
        alert('无法初始化裁剪器，请换一张图片重试。');
        closeCropModal();
      }
    }

    cropTarget.onerror = function () {
      cropTarget.onerror = null;
      cropTarget.onload = null;
      alert('图片无法加载，无法裁剪。');
      closeCropModal();
    };
    cropTarget.onload = function () {
      cropTarget.onload = null;
      cropTarget.onerror = null;
      requestAnimationFrame(startCropper);
    };
    cropTarget.src = src;
  }

  function confirmCrop() {
    if (!cropperInstance || cropRowIndex < 0 || cropRowIndex >= tableData.length) {
      closeCropModal();
      return;
    }
    var canvas = null;
    try {
      canvas = cropperInstance.getCroppedCanvas({
        maxWidth: CROP_MAX_EDGE,
        maxHeight: CROP_MAX_EDGE,
        imageSmoothingEnabled: true,
        imageSmoothingQuality: 'high',
      });
    } catch (e) {
      console.error(e);
    }
    if (!canvas || canvas.width === 0) {
      alert('无法生成裁剪图（部分格式如 SVG 可能不支持）。');
      return;
    }
    var dataUrl;
    try {
      dataUrl = canvas.toDataURL('image/jpeg', 0.92);
    } catch (e) {
      console.error(e);
      alert('导出裁剪图失败，请尝试更换图片格式。');
      return;
    }
    tableData[cropRowIndex].imgSrc = dataUrl;
    saveToLocalStorage();
    closeCropModal();
    renderTable();
  }

  function readFileAsDataURL(file, callback) {
    var reader = new FileReader();
    reader.onload = function () {
      callback(reader.result);
    };
    reader.onerror = function () {
      callback(null);
    };
    reader.readAsDataURL(file);
  }

  function onCharacterImageSelected(ev) {
    var file = ev.target && ev.target.files && ev.target.files[0];
    if (!file || !file.type || file.type.indexOf('image/') !== 0) {
      pendingUploadRowIndex = -1;
      return;
    }
    readFileAsDataURL(file, function (dataUrl) {
      if (dataUrl && pendingUploadRowIndex >= 0 && pendingUploadRowIndex < tableData.length) {
        openCropModal(pendingUploadRowIndex, dataUrl);
      }
      pendingUploadRowIndex = -1;
    });
  }

  function onBgImageSelected(ev) {
    var file = ev.target && ev.target.files && ev.target.files[0];
    if (!file || !file.type || file.type.indexOf('image/') !== 0) return;
    readFileAsDataURL(file, function (dataUrl) {
      if (dataUrl) {
        meta.bgImageSrc = dataUrl;
        saveToLocalStorage();
        applyPreviewStyles();
      }
    });
  }

  function ensureSelectionInCell(cell) {
    if (!cell) return false;
    var sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return false;
    return cell.contains(sel.anchorNode);
  }

  function captureSelectionInCurrentCell() {
    if (!currentRichCell) return;
    var sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    try {
      var r = sel.getRangeAt(0);
      if (currentRichCell.contains(r.commonAncestorContainer)) {
        savedSelectionRange = r.cloneRange();
      }
    } catch (e) {
      savedSelectionRange = null;
    }
  }

  function restoreSelectionInCell() {
    if (!savedSelectionRange || !currentRichCell) return false;
    try {
      currentRichCell.focus();
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedSelectionRange);
      return true;
    } catch (e) {
      savedSelectionRange = null;
      return false;
    }
  }

  function preventToolbarFocusSteal(el) {
    if (!el) return;
    el.addEventListener('mousedown', function (e) {
      e.preventDefault();
    });
  }

  function focusCellEnd(cell) {
    cell.focus();
    var range = document.createRange();
    range.selectNodeContents(cell);
    range.collapse(false);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function execOnRichCell(command, value) {
    if (!currentRichCell || !document.body.contains(currentRichCell)) {
      alert('请先在左侧预览中点击「初印象」单元格。');
      return false;
    }
    currentRichCell.focus();
    restoreSelectionInCell();
    if (!ensureSelectionInCell(currentRichCell)) {
      focusCellEnd(currentRichCell);
    }
    try {
      document.execCommand('styleWithCSS', false, true);
      return document.execCommand(command, false, value != null ? value : null);
    } catch (e) {
      console.warn('execCommand 失败:', command, e);
      return false;
    } finally {
      var idx = parseInt(currentRichCell.dataset.rowIndex, 10);
      if (!isNaN(idx)) syncRichTextFromCell(idx, currentRichCell);
    }
  }

  function getFontSizeAtSelectionStart() {
    var sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return 16;
    var node = sel.anchorNode;
    var el = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
    if (!el || !document.body.contains(el)) return 16;
    var px = parseFloat(window.getComputedStyle(el).fontSize);
    return isNaN(px) ? 16 : px;
  }

  function adjustFontSize(delta) {
    if (!currentRichCell || !document.body.contains(currentRichCell)) {
      alert('请先在左侧预览中点击「初印象」单元格。');
      return;
    }
    currentRichCell.focus();
    restoreSelectionInCell();
    var sel = window.getSelection();
    var hasRange = sel && sel.rangeCount > 0 && !sel.getRangeAt(0).collapsed;

    if (hasRange && ensureSelectionInCell(currentRichCell)) {
      var range = sel.getRangeAt(0);
      var base = getFontSizeAtSelectionStart();
      var next = Math.min(40, Math.max(10, Math.round(base + delta)));
      try {
        var span = document.createElement('span');
        span.style.fontSize = next + 'px';
        var contents = range.extractContents();
        span.appendChild(contents);
        range.insertNode(span);
        sel.removeAllRanges();
        var nr = document.createRange();
        nr.selectNodeContents(span);
        nr.collapse(false);
        sel.addRange(nr);
      } catch (e) {
        adjustWholeCellFontSize(currentRichCell, delta);
      }
    } else {
      adjustWholeCellFontSize(currentRichCell, delta);
    }

    var idx = parseInt(currentRichCell.dataset.rowIndex, 10);
    if (!isNaN(idx)) syncRichTextFromCell(idx, currentRichCell);
  }

  function adjustWholeCellFontSize(cell, delta) {
    var cs = window.getComputedStyle(cell);
    var px = parseFloat(cs.fontSize) || 16;
    var next = Math.min(40, Math.max(10, Math.round(px + delta)));
    cell.style.fontSize = next + 'px';
  }

  function applyForeColor(hex) {
    if (!currentRichCell || !document.body.contains(currentRichCell)) {
      alert('请先在左侧预览中点击「初印象」单元格。');
      return;
    }
    currentRichCell.focus();
    restoreSelectionInCell();
    if (!ensureSelectionInCell(currentRichCell)) {
      focusCellEnd(currentRichCell);
    }
    try {
      document.execCommand('styleWithCSS', false, true);
      document.execCommand('foreColor', false, hex);
    } catch (e) {
      console.warn('foreColor 失败:', e);
    }
    var idx = parseInt(currentRichCell.dataset.rowIndex, 10);
    if (!isNaN(idx)) syncRichTextFromCell(idx, currentRichCell);
  }

  function addRow() {
    tableData.push(
      normalizeRow({
        imgSrc: '',
        richFirst: '<p>新角色</p>',
        richCurrent: '<p></p>',
        rowMinHeight: 0,
      })
    );
    saveToLocalStorage();
    renderTable();
  }

  function removeLastRow() {
    if (tableData.length <= 1) return;
    tableData.pop();
    saveToLocalStorage();
    renderTable();
  }

  function downloadBlob(blob, filename, format) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.download = filename;
    link.href = url;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    
    document.body.appendChild(link);
    link.click();
    
    setTimeout(function() {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 100);
  }

  // 检测是否为移动设备
  function isMobileDevice() {
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
  }

  async function exportImage() {
    const captureEl = document.getElementById('captureArea');
    if (!captureEl) {
      console.error('找不到截图元素 #captureArea');
      alert('导出失败：找不到预览区域');
      return;
    }

    const format = document.getElementById('exportFormat').value;
    const presetId = document.getElementById('exportPreset').value;
    const preset = EXPORT_PRESETS[presetId] || EXPORT_PRESETS['xhs_3_4'];

    // 临时创建一个专门用于导出的容器，避免样式冲突
    const exportContainer = document.createElement('div');
    exportContainer.id = 'temp-export-container';
    
    // 对于固定高度的预设，使用预设高度；对于自适应高度的预设，初始高度设为较大值，后面再调整
    const containerHeight = (preset.mode === 'fitWidth' || preset.h == null) ? Math.max(2000, captureEl.scrollHeight) : preset.h;
    
    exportContainer.style.cssText = `
      position: absolute;
      left: 0;
      top: 0;
      width: ${preset.w}px;
      min-height: ${containerHeight}px;
      background: #ffffff;
      padding: 28px 36px 40px;
      z-index: 9999;
      visibility: hidden;
      pointer-events: none;
      font-family: "Segoe UI", system-ui, -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
      font-size: 16px;
      line-height: 1.5;
      color: #1e293b;
      box-sizing: border-box;
    `;

    // 克隆内容
    const clonedContent = captureEl.cloneNode(true);
    clonedContent.id = 'cloned-for-export';
    clonedContent.style.cssText = `
      width: 100%;
      min-height: auto;
      background: #ffffff;
      position: relative;
      margin: 0;
      visibility: visible;
      opacity: 1;
      box-shadow: none;
      transform: none;
      display: block;
      max-width: none;
      max-height: none;
      overflow: visible;
      font-family: inherit;
      font-size: inherit;
      line-height: inherit;
      color: inherit;
    `;

    // 处理需要隐藏的元素
    const hideEls = clonedContent.querySelectorAll('.hide-on-export, .image-controls, .row-resize-handle, .col-resize-handle');
    hideEls.forEach(el => {
      el.style.display = 'none';
    });

    // 确保所有内容可见
    const allElements = clonedContent.querySelectorAll('*');
    allElements.forEach(el => {
      if (el.style.display === 'none') return;
      if (el.tagName === 'IMG' && !el.src) {
        el.style.display = 'none';
      }
    });

    // 特别处理表格中的富文本内容
    const richCells = clonedContent.querySelectorAll('.rich-cell');
    richCells.forEach(cell => {
      // 确保富文本内容可见
      if (!cell.innerHTML || cell.innerHTML.trim() === '') {
        cell.innerHTML = '<p>（无内容）</p>';
      }
      // 保留内容但隐藏编辑相关的属性
      cell.contentEditable = 'false';
      cell.removeAttribute('data-placeholder');
    });

    exportContainer.appendChild(clonedContent);
    document.body.appendChild(exportContainer);

    // 等待内容渲染
    await new Promise(resolve => setTimeout(resolve, 100));

    try {
      // 使用 html2canvas 截图克隆的内容
      const canvas = await html2canvas(exportContainer, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
        allowTaint: true,
        logging: false,
        width: exportContainer.offsetWidth,
        height: exportContainer.scrollHeight,  // 使用实际滚动高度
        foreignObjectRendering: true,
        ignoreElements: (element) => {
          return element.classList.contains('hide-on-export');
        },
        // 确保所有元素都正确渲染
        onclone: (clonedDoc) => {
          // 在克隆文档中确保所有元素可见
          const allClonedElements = clonedDoc.querySelectorAll('*');
          allClonedElements.forEach(el => {
            if (el.style && el.style.display === 'none' && 
                (el.classList.contains('hide-on-export') || 
                 el.classList.contains('image-controls') ||
                 el.classList.contains('row-resize-handle'))) {
              // 这些元素应该保持隐藏
            } else if (el.style) {
              // 确保其他元素可见
              if (el.style.display === 'none') el.style.display = 'block';
              if (el.style.visibility === 'hidden') el.style.visibility = 'visible';
              if (el.style.opacity === '0') el.style.opacity = '1';
            }
          });
        }
      });

      // 根据预设尺寸缩放画布
      const finalCanvas = resizeCanvasToExport(canvas, preset, '#ffffff');

      // 生成图片
      const blob = await new Promise(resolve => {
        const mime = format === 'jpeg' ? 'image/jpeg' : 'image/png';
        const quality = format === 'jpeg' ? 0.92 : 0.95;
        finalCanvas.toBlob(resolve, mime, quality);
      });

      if (blob && blob.size > 0) {
        // 创建对象URL
        const imageUrl = URL.createObjectURL(blob);
        const suffix = preset.mode === 'fitWidth' ? preset.w + 'w' : preset.w + 'x' + preset.h;
        const filename = `角色印象表_${presetId}_${suffix}.${format === 'jpeg' ? 'jpg' : 'png'}`;
        
        // 检测是否为移动设备
        if (isMobileDevice()) {
          // 移动端：显示导出的图片，提示用户长按保存
          exportedImageDisplay.src = imageUrl;
          exportedImageModal.classList.add('is-open');
          exportedImageModal.setAttribute('aria-hidden', 'false');
        } else {
          // 桌面端：继续原来的下载方式
          downloadBlob(blob, filename, format);
        }
      } else {
        alert('导出失败：生成的图片为空');
      }
    } catch (error) {
      console.error('导出错误:', error);
      alert('导出失败：' + error.message);
    } finally {
      // 清理临时容器
      if (exportContainer.parentNode) {
        exportContainer.parentNode.removeChild(exportContainer);
      }
    }
  }

  function bindMetaInputs() {
    if (workTitle) {
      workTitle.addEventListener('input', function () {
        meta.workTitle = workTitle.value;
        saveToLocalStorage();
      });
    }
    if (filledBy) {
      filledBy.addEventListener('input', function () {
        meta.filledBy = filledBy.value;
        saveToLocalStorage();
      });
    }
    if (madeBy) {
      madeBy.addEventListener('input', function () {
        meta.madeBy = madeBy.value;
        saveToLocalStorage();
      });
    }
    if (bgBaseTone) {
      bgBaseTone.addEventListener('change', function () {
        meta.bgBaseTone = bgBaseTone.value;
        saveToLocalStorage();
        applyPreviewStyles();
      });
    }
    if (bgOpacity) {
      bgOpacity.addEventListener('input', function () {
        var v = parseInt(bgOpacity.value, 10);
        if (isNaN(v)) return;
        meta.bgOpacity = v;
        if (bgOpacityVal) bgOpacityVal.textContent = v + '%';
        if (previewBgLayer) previewBgLayer.style.opacity = String(v / 100);
        saveToLocalStorage();
      });
    }
    if (btnBgImage && hiddenBgInput) {
      btnBgImage.addEventListener('click', function () {
        hiddenBgInput.value = '';
        hiddenBgInput.click();
      });
    }
    if (hiddenBgInput) {
      hiddenBgInput.addEventListener('change', onBgImageSelected);
    }
    if (btnClearBg) {
      btnClearBg.addEventListener('click', function () {
        meta.bgImageSrc = '';
        saveToLocalStorage();
        applyPreviewStyles();
      });
    }
    if (exportPreset) {
      exportPreset.addEventListener('change', function () {
        meta.exportPreset = exportPreset.value;
        if (!EXPORT_PRESETS[meta.exportPreset]) meta.exportPreset = 'xhs_3_4';
        saveToLocalStorage();
        applyExportPresetFrame();
      });
    }
  }

  function bindImageModals() {
    if (imageActionModal) {
      imageActionModal.addEventListener('click', function (ev) {
        if (ev.target === imageActionModal) closeImageActionModal();
      });
    }
    if (btnImageReplace) {
      btnImageReplace.addEventListener('click', function () {
        var r = pendingImageActionRowIndex;
        closeImageActionModal();
        if (r >= 0) openFilePicker(r);
      });
    }
    if (btnImageCrop) {
      btnImageCrop.addEventListener('click', function () {
        var r = pendingImageActionRowIndex;
        closeImageActionModal();
        if (r >= 0) openCropModal(r);
      });
    }
    if (btnImageActionCancel) {
      btnImageActionCancel.addEventListener('click', closeImageActionModal);
    }
    if (cropModal) {
      cropModal.addEventListener('click', function (ev) {
        if (ev.target === cropModal) closeCropModal();
      });
    }
    if (btnCropCancel) btnCropCancel.addEventListener('click', closeCropModal);
    if (btnCropConfirm) btnCropConfirm.addEventListener('click', confirmCrop);
    
    // 绑定导出图片模态框事件
    if (exportedImageModal) {
      exportedImageModal.addEventListener('click', function (ev) {
        if (ev.target === exportedImageModal) {
          exportedImageModal.classList.remove('is-open');
          exportedImageModal.setAttribute('aria-hidden', 'true');
        }
      });
    }
    if (btnCloseExportModal) {
      btnCloseExportModal.addEventListener('click', function() {
        exportedImageModal.classList.remove('is-open');
        exportedImageModal.setAttribute('aria-hidden', 'true');
      });
    }

    document.addEventListener('keydown', function (ev) {
      if (ev.key !== 'Escape') return;
      if (cropModal && cropModal.classList.contains('is-open')) {
        ev.preventDefault();
        closeCropModal();
      } else if (imageActionModal && imageActionModal.classList.contains('is-open')) {
        ev.preventDefault();
        closeImageActionModal();
      } else if (exportedImageModal && exportedImageModal.classList.contains('is-open')) {
        ev.preventDefault();
        exportedImageModal.classList.remove('is-open');
        exportedImageModal.setAttribute('aria-hidden', 'true');
      }
    });
  }

  function onColResizeMove(e) {
    if (!colDrag) return;
    var d = ((e.clientX - colDrag.startX) / colDrag.tableW) * 100;
    meta.imageColWidthPct =
      Math.round(Math.min(62, Math.max(22, colDrag.startPct + d)) * 10) / 10;
    applyImageColPct();
  }

  function onColResizeEnd() {
    document.removeEventListener('mousemove', onColResizeMove);
    document.removeEventListener('mouseup', onColResizeEnd);
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
    if (colDrag) {
      colDrag = null;
      saveToLocalStorage();
    }
  }

  function bindColResize() {
    if (!colResizeHandle || !mainTable) return;
    colResizeHandle.addEventListener('mousedown', function (e) {
      e.preventDefault();
      e.stopPropagation();
      var rect = mainTable.getBoundingClientRect();
      colDrag = {
        startX: e.clientX,
        startPct: meta.imageColWidthPct,
        tableW: Math.max(120, rect.width),
      };
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';
      document.addEventListener('mousemove', onColResizeMove);
      document.addEventListener('mouseup', onColResizeEnd);
    });
  }

  function onRowResizeMove(e) {
    if (!rowDrag) return;
    var dy = e.clientY - rowDrag.startY;
    var next = Math.round(Math.max(72, rowDrag.startMin + dy));
    tableData[rowDrag.rowIndex].rowMinHeight = next;
    var tr = tableBody && tableBody.querySelector('tr[data-row-index="' + rowDrag.rowIndex + '"]');
    if (tr) {
      tr.style.minHeight = next + 'px';
      var tds = tr.querySelectorAll('td');
      for (var i = 0; i < tds.length; i++) {
        tds[i].style.minHeight = next + 'px';
      }
    }
  }

  function onRowResizeEnd() {
    document.removeEventListener('mousemove', onRowResizeMove);
    document.removeEventListener('mouseup', onRowResizeEnd);
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
    if (rowDrag) {
      rowDrag = null;
      saveToLocalStorage();
    }
  }

  function bindRowResize() {
    if (!tableBody) return;
    tableBody.addEventListener('mousedown', function (e) {
      var h = e.target.closest('.row-resize-handle');
      if (!h || !tableBody.contains(h)) return;
      e.preventDefault();
      e.stopPropagation();
      var ri = parseInt(h.dataset.rowIndex, 10);
      if (isNaN(ri) || ri < 0 || ri >= tableData.length) return;
      var tr = h.closest('tr');
      var rect = tr ? tr.getBoundingClientRect() : null;
      var base =
        tableData[ri].rowMinHeight > 0
          ? tableData[ri].rowMinHeight
          : rect
            ? Math.round(rect.height)
            : 100;
      rowDrag = { rowIndex: ri, startY: e.clientY, startMin: base };
      document.body.style.cursor = 'ns-resize';
      document.body.style.userSelect = 'none';
      document.addEventListener('mousemove', onRowResizeMove);
      document.addEventListener('mouseup', onRowResizeEnd);
    });
  }

  function bindEvents() {
    bindMetaInputs();
    bindImageModals();
    bindColResize();
    bindRowResize();
    if (btnAddRow) btnAddRow.addEventListener('click', addRow);
    if (btnRemoveRow) btnRemoveRow.addEventListener('click', removeLastRow);
    if (btnExport) btnExport.addEventListener('click', exportImage);
    if (hiddenFileInput) hiddenFileInput.addEventListener('change', onCharacterImageSelected);

    if (btnBold) btnBold.addEventListener('click', function () { execOnRichCell('bold'); });
    if (btnItalic) btnItalic.addEventListener('click', function () { execOnRichCell('italic'); });
    if (btnUnderline) btnUnderline.addEventListener('click', function () { execOnRichCell('underline'); });
    if (btnFontInc) btnFontInc.addEventListener('click', function () { adjustFontSize(2); });
    if (btnFontDec) btnFontDec.addEventListener('click', function () { adjustFontSize(-2); });
    if (colorPicker) {
      colorPicker.addEventListener('mousedown', captureSelectionInCurrentCell);
      colorPicker.addEventListener('input', function () {
        applyForeColor(colorPicker.value);
      });
      colorPicker.addEventListener('change', function () {
        applyForeColor(colorPicker.value);
      });
    }

    preventToolbarFocusSteal(btnBold);
    preventToolbarFocusSteal(btnItalic);
    preventToolbarFocusSteal(btnUnderline);
    preventToolbarFocusSteal(btnFontInc);
    preventToolbarFocusSteal(btnFontDec);
  }

  function init() {
    var loaded = loadFromLocalStorage() || loadLegacyV1();
    if (loaded) {
      meta = normalizeMeta(loaded.meta || {});
      tableData = loaded.rows;
    } else {
      meta = getDefaultMeta();
      tableData = defaultRows();
    }
    pushMetaToInputs();
    applyPreviewStyles();
    applyExportPresetFrame();
    bindEvents();
    renderTable();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();