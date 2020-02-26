var UseListSheet = function(targetName) {
  this.divName = targetName;
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('D利用(' + targetName + ')');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  
  this.showTitleError = function(key) {
    Browser.msgBox(ERROR_TEXT_NONKEY_TITLE, ERROR_TEXT_NONKEY_MESSAGE + key, Browser.Buttons.OK);
  }
}
  
UseListSheet.prototype = {
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
    var peopleSheetIndex = peopleSheet.getIndex();
    var divName = this.divName
    var target = peopleSheet.values.filter(function(value) { return value[peopleSheetIndex.div] === divName });
    var titles = KintonePCData.pcDataSheet.getTitles();
    var pcData = [];
    
    var data = target.filter(function(value) {
      var hasPc = KintonePCData.pcDataSheet.getTargetUserData(value[peopleSheetIndex.cynumber]);
      var isTarget = (hasPc.length === 1 && hasPc[0][titles.pc_category.index] === 'D');
      if (isTarget) pcData.push([
        hasPc[0][titles.capc_id.index],
        hasPc[0][titles.pc_maker.index],
        hasPc[0][titles.pc_product.index],
        hasPc[0][titles.pc_model.index],
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
    sheet.getRange('I1').setValue(data.length + ' / ' + target.length + '人');
  }
};

function updateDListForCa() {
  var useListHonbuSheet    = new UseListSheet('ｲﾝﾀｰﾈｯﾄ広告事業本部');
  var useListCaDesignSheet = new UseListSheet('株式会社Ca Design');
  var useListAiSheet       = new UseListSheet('AI事業本部');
  useListHonbuSheet.updateData();
  useListCaDesignSheet.updateData();
  useListAiSheet.updateData();
}
