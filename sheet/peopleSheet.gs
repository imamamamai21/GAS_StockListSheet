var PeopleSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('人員データ(Kintone)');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  
  this.showTitleError = function(key) {
    Browser.msgBox(ERROR_TEXT_NONKEY_TITLE, ERROR_TEXT_NONKEY_MESSAGE + key, Browser.Buttons.OK);
  }
}
  
PeopleSheet.prototype = {
  getRowKey: function(target) {
    var index = this.getIndex();
    var targetIndex = index[target];
    var returnKey = (targetIndex > -1) ? SHEET_ROWS[targetIndex] : '';
    if (!returnKey || returnKey === '') this.showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  createIndex: function() {
    const NO = '社員番号';
    var filterData = this.values.filter(function(value) {
      return value.indexOf(NO) > -1;
    })[0];
    if(!filterData || filterData.length === 0) {
      this.showTitleError();
      return;
    }
    
    this.index = {
      cynumber   : filterData.indexOf(NO),
      name       : filterData.indexOf('名前'),
      verbose_name: filterData.indexOf('よみがな'),
      mail       : filterData.indexOf('メール'),
      tel        : filterData.indexOf('電話'),
      company    : filterData.indexOf('会社'),
      div        : filterData.indexOf('部署'),
      div1       : filterData.indexOf('部署1'),
      div2       : filterData.indexOf('部署2'),
      div3       : filterData.indexOf('部署3'),
      join_date  : filterData.indexOf('入社日')
    };
    return this.index;
  },
  updateData: function() {
    var repsponse = KintoneApi.manMasterApi.getAllData();
    var sortObj = {};
    var index = this.getIndex();
    var sheet = this.sheet;
    
    // 項目ごとのobjectに変換する
    repsponse.forEach(function(values, i) {
      Object.keys(values).forEach(function(key) {
        if (key === 'レコード番号') return;
        if (!sortObj[key]) sortObj[key] = [];
        sortObj[key][i] = [values[key].value];
      });
    });
    Object.keys(sortObj).forEach(function (key) { // 項目ごとに書き込む
      sheet.getRange(3, index[key] + 1, sortObj[key].length, 1).setValues(sortObj[key]);
    });
    // 更新日時のアップデート
    var timeStamp = Utilities.formatDate(new Date(), 'JST', 'yyyy年 MM/dd(E) HH:mm');
    sheet.getRange('E1').setValue(timeStamp);
  }
};
var peopleSheet = new PeopleSheet();

function updatePeple() {
  peopleSheet.updateData();
}
