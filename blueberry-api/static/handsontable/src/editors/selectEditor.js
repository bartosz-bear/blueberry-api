
import * as dom from './../dom.js';
import * as helper from './../helpers.js';
import {getEditor, registerEditor} from './../editors.js';
import {BaseEditor} from './_baseEditor.js';

var SelectEditor = BaseEditor.prototype.extend();


/**
 * @private
 * @editor SelectEditor
 * @class SelectEditor
 */
SelectEditor.prototype.init = function() {
  this.select = document.createElement('SELECT');
  dom.addClass(this.select, 'htSelectEditor');
  this.select.style.display = 'none';
  this.instance.rootElement.appendChild(this.select);
  this.registerHooks();
};

SelectEditor.prototype.registerHooks = function() {
  this.instance.addHook('afterScrollVertically', () => this.refreshDimensions());
  this.instance.addHook('afterColumnResize', () => this.refreshDimensions());
  this.instance.addHook('afterRowResize', () => this.refreshDimensions());
};

SelectEditor.prototype.prepare = function() {
  BaseEditor.prototype.prepare.apply(this, arguments);

  var selectOptions = this.cellProperties.selectOptions;
  var options;

  if (typeof selectOptions == 'function') {
    options = this.prepareOptions(selectOptions(this.row, this.col, this.prop));
  } else {
    options = this.prepareOptions(selectOptions);
  }

  dom.empty(this.select);

  for (var option in options) {
    if (options.hasOwnProperty(option)) {
      var optionElement = document.createElement('OPTION');
      optionElement.value = option;
      dom.fastInnerHTML(optionElement, options[option]);
      this.select.appendChild(optionElement);
    }
  }
};

SelectEditor.prototype.prepareOptions = function(optionsToPrepare) {
  var preparedOptions = {};

  if (Array.isArray(optionsToPrepare)) {
    for (var i = 0, len = optionsToPrepare.length; i < len; i++) {
      preparedOptions[optionsToPrepare[i]] = optionsToPrepare[i];
    }
  } else if (typeof optionsToPrepare == 'object') {
    preparedOptions = optionsToPrepare;
  }

  return preparedOptions;

};

SelectEditor.prototype.getValue = function() {
  return this.select.value;
};

SelectEditor.prototype.setValue = function(value) {
  this.select.value = value;
};

var onBeforeKeyDown = function(event) {
  var instance = this;
  var editor = instance.getActiveEditor();

  if (event != null && event.isImmediatePropagationEnabled == null) {
    event.stopImmediatePropagation = function() {
      this.isImmediatePropagationEnabled = false;
    };
    event.isImmediatePropagationEnabled = true;
    event.isImmediatePropagationStopped = function() {
      return !this.isImmediatePropagationEnabled;
    };
  }

  switch (event.keyCode) {
    case helper.keyCode.ARROW_UP:
      var previousOptionIndex = editor.select.selectedIndex - 1;
      if (previousOptionIndex >= 0) {
        editor.select[previousOptionIndex].selected = true;
      }

      event.stopImmediatePropagation();
      event.preventDefault();
      break;

    case helper.keyCode.ARROW_DOWN:
      var nextOptionIndex = editor.select.selectedIndex + 1;
      if (nextOptionIndex <= editor.select.length - 1) {
        editor.select[nextOptionIndex].selected = true;
      }

      event.stopImmediatePropagation();
      event.preventDefault();
      break;
  }
};

// TODO: Refactor this with the use of new getCell() after 0.12.1
SelectEditor.prototype.checkEditorSection = function() {
  if (this.row < this.instance.getSettings().fixedRowsTop) {
    if (this.col < this.instance.getSettings().fixedColumnsLeft) {
      return 'corner';
    } else {
      return 'top';
    }
  } else {
    if (this.col < this.instance.getSettings().fixedColumnsLeft) {
      return 'left';
    }
  }
};

SelectEditor.prototype.open = function() {
  this._opened = true;
  this.refreshDimensions();
  this.select.style.display = '';
  this.instance.addHook('beforeKeyDown', onBeforeKeyDown);
};

SelectEditor.prototype.close = function() {
  this._opened = false;
  this.select.style.display = 'none';
  this.instance.removeHook('beforeKeyDown', onBeforeKeyDown);
};

SelectEditor.prototype.focus = function() {
  this.select.focus();
};

SelectEditor.prototype.refreshDimensions = function() {
  if (this.state !== Handsontable.EditorState.EDITING) {
    return;
  }
  this.TD = this.getEditedCell();

  // TD is outside of the viewport.
  if (!this.TD) {
    this.close();

    return;
  }
  var
    width = dom.outerWidth(this.TD) + 1,
    height = dom.outerHeight(this.TD) + 1,
    currentOffset = dom.offset(this.TD),
    containerOffset = dom.offset(this.instance.rootElement),
    scrollableContainer = dom.getScrollableElement(this.TD),
    editTop = currentOffset.top - containerOffset.top - 1 - (scrollableContainer.scrollTop || 0),
    editLeft = currentOffset.left - containerOffset.left - 1 - (scrollableContainer.scrollLeft || 0),
    editorSection = this.checkEditorSection(),
    cssTransformOffset;

  const settings = this.instance.getSettings();
  let rowHeadersCount = settings.rowHeaders ? 1 : 0;
  let colHeadersCount = settings.colHeaders ? 1 : 0;

  switch (editorSection) {
    case 'top':
      cssTransformOffset = dom.getCssTransform(this.instance.view.wt.wtOverlays.topOverlay.clone.wtTable.holder.parentNode);
      break;
    case 'left':
      cssTransformOffset = dom.getCssTransform(this.instance.view.wt.wtOverlays.leftOverlay.clone.wtTable.holder.parentNode);
      break;
    case 'corner':
      cssTransformOffset = dom.getCssTransform(this.instance.view.wt.wtOverlays.topLeftCornerOverlay.clone.wtTable.holder.parentNode);
      break;
  }
  if (this.instance.getSelected()[0] === 0) {
    editTop += 1;
  }

  if (this.instance.getSelected()[1] === 0) {
    editLeft += 1;
  }

  var selectStyle = this.select.style;

  if (cssTransformOffset && cssTransformOffset != -1) {
    selectStyle[cssTransformOffset[0]] = cssTransformOffset[1];
  } else {
    dom.resetCssTransform(this.select);
  }
  const cellComputedStyle = dom.getComputedStyle(this.TD);

  if (parseInt(cellComputedStyle.borderTopWidth, 10) > 0) {
    height -= 1;
  }
  if (parseInt(cellComputedStyle.borderLeftWidth, 10) > 0) {
    width -= 1;
  }

  selectStyle.height = height + 'px';
  selectStyle.minWidth = width + 'px';
  selectStyle.top = editTop + 'px';
  selectStyle.left = editLeft + 'px';
  selectStyle.margin = '0px';
};

SelectEditor.prototype.getEditedCell = function() {
  var editorSection = this.checkEditorSection(),
    editedCell;

  switch (editorSection) {
    case 'top':
      editedCell = this.instance.view.wt.wtOverlays.topOverlay.clone.wtTable.getCell({
        row: this.row,
        col: this.col
      });
      this.select.style.zIndex = 101;
      break;
    case 'corner':
      editedCell = this.instance.view.wt.wtOverlays.topLeftCornerOverlay.clone.wtTable.getCell({
        row: this.row,
        col: this.col
      });
      this.select.style.zIndex = 103;
      break;
    case 'left':
      editedCell = this.instance.view.wt.wtOverlays.leftOverlay.clone.wtTable.getCell({
        row: this.row,
        col: this.col
      });
      this.select.style.zIndex = 102;
      break;
    default:
      editedCell = this.instance.getCell(this.row, this.col);
      this.select.style.zIndex = '';
      break;
  }

  return editedCell != -1 && editedCell != -2 ? editedCell : void 0;
};

export {SelectEditor};

registerEditor('select', SelectEditor);

