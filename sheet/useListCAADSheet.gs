var UseListCAADSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('D利用(株式会社CAAD)');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  
  this.showTitleError = function(key) {
    Browser.msgBox(ERROR_TEXT_NONKEY_TITLE, ERROR_TEXT_NONKEY_MESSAGE + key, Browser.Buttons.OK);
  }
}
  
UseListCAADSheet.prototype = {
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
    const NO = 'PC番号';
    var filterData = this.values.filter(function(value) {
      return value.indexOf(NO) > -1;
    })[0];
    if(!filterData || filterData.length === 0) {
      this.showTitleError();
      return;
    }
    
    this.index = {
      pcNo       : filterData.indexOf(NO),
      maker      : filterData.indexOf('メーカー'),
      product    : filterData.indexOf('製品名'),
      model      : filterData.indexOf('モデル'),
      cynumber   : filterData.indexOf('社員番号'),
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
    var targetPc = KintoneApi.caadApi.getInUseData();
    var peopleSheetIndex = peopleSheet.getIndex();
    var targetUser = peopleSheet.values.filter(function(value) { return value[peopleSheetIndex.company] === '株式会社ｼｰｴｰ･ｱﾄﾞﾊﾞﾝｽ' });
    var pcData = [];
    
    var data = targetUser.filter(function(value) {
      var hasPc = targetPc.filter(function(pcValue) {
        return pcValue.user_id.value === value[peopleSheetIndex.cynumber];
      });
      var isTarget = (hasPc.length === 1 && hasPc[0].pc_category.value === 'D');
      if (isTarget) pcData.push([
        hasPc[0].capc_id.value,
        hasPc[0].pc_maker.value,
        hasPc[0].pc_product.value,
        hasPc[0].pc_model.value,
      ]);
      return isTarget;
    })

    var index = this.getIndex();
    var sheet = this.sheet;
    sheet.getRange(3, 1, sheet.getLastRow(), Object.keys(index).length - 1).clearContent();
    
    data.forEach(function(value, i) {
      sheet.getRange(3 + i, 1, 1, pcData[i].length).setValues([pcData[i]]);
      sheet.getRange(3 + i, 1 + pcData[i].length, 1, value.length).setValues([value]);
    })
    
    // 更新日時のアップデート
    var timeStamp = Utilities.formatDate(new Date(), 'JST', 'yyyy年 MM/dd(E) HH:mm');
    sheet.getRange('G1').setValue(timeStamp);
    sheet.getRange('I1').setValue(data.length + ' / ' + targetUser.length + '人');
  }
};

function updateDListForCaad() {
  var useListCAADSheet = new UseListCAADSheet();
  useListCAADSheet.updateData();
}
