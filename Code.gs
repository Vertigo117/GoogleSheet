class ModalDialog {
  constructor() {
    this._name = "Form1";
    this._width = 300;
    this._height = 300;
  }

  get name() {
    return this._name;
  }

  set name(value) {
    this._name = value;
  }

  get width() {
    return this._width;
  }

  set width(value) {
    this._width = value;
  }

  get height() {
    return this._height;
  }

  set height(value) {
    this._height = value;
  }

  showFromHtml(filename) {
    let html = HtmlService.createHtmlOutputFromFile(filename).setWidth(this._width).setHeight(this._height);
    SpreadsheetApp.getUi().showModalDialog(html, this._name);
  }
}

function onOpen() {
  Run();
}

function Run() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Run')
      .addItem('Run', 'campDup')
      .addItem('Duplicate Adset', 'showDuplicateAdsetDialog')
      .addItem('Adjust Bid Amount and Budget', 'showAdjustDialog')
      .addToUi();
}

function showDuplicateAdsetDialog() {
  let modalDialog = new ModalDialog();  
  modalDialog.name = 'New Adset';
  modalDialog.width = 400;
  modalDialog.height = 400;
  modalDialog.showFromHtml('ModalDialog');
}

function showAdjustDialog() {
  let modalDialog = new ModalDialog();
  modalDialog.name = 'Adjust Bid Amount and Budget';
  modalDialog.width = 1000;
  modalDialog.height = 600;
  modalDialog.showFromHtml('AdjustBudgetDialog');
}

function getAdsetsForDisplay() {
  let adsets = getColumnValuesArray('Ad Set Name');
  let adsetsDistinct = adsets.filter((value, index, adsets) => adsets.indexOf(value) === index);

  return adsetsDistinct;
}

function applyAdjustments(objectArr) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let rangeData = sheet.getDataRange();
  let lastColumn = rangeData.getLastColumn();
  let lastRow = rangeData.getLastRow();
  let columnNames = [];

  for(let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
    let columnName = sheet.getRange(1, columnIndex).getValue();
    columnNames.push(columnName);
  }

  for(let obj of objectArr) {
    for(let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
      let adsetName = sheet.getRange(rowIndex, columnNames.indexOf('Ad Set Name') + 1).getValue();

      if(obj['Ad Set Name'] == adsetName) {
        sheet.getRange(rowIndex, columnNames.indexOf('Bid Amount') + 1).setValue(obj['Bid Amount'])
        sheet.getRange(rowIndex, columnNames.indexOf('Ad Set Daily Budget') + 1).setValue(obj['Ad Set Daily Budget']);
      }
    }
  }

  SpreadsheetApp.getUi().alert('Operation Complete!');
}

function createNewAdset(versions) {
  //versions = ['JR - Bloopers V5 Auto CAUS S1 UV06 III A7'];
  //let spreadsheet = new SpreadSheet(SpreadsheetApp);
  let startPoint = 2;

  var geos = ['US', 'CA', 'AU', 'WW', 'CAUS' , 'CAAUUS','NAC','CAU'];
    var platforms = ['Mob', 'Desk', 'Auto','AutoAnd','MobA','IOS'];
    var sites = ['TF','JR','BB'];
    var accounts = ['I', 'II', 'III', 'IV', 'V', 'TF1', 'TF2', 'TF3', '6', '10', '11', '12', '14', '15', '16', '17', '18', '19', '20'];
    var uvs = ['UV01', 'UV02', 'UV03', 'UV04', 'UV05', 'UV06', 'UV07', 'UV08', 'UV09', 'UV10', 'UV11', 'UV12', 'UV13', 'UV14', 'UV15','UVG04','UVG03','UV04F'];
    var uvMap = {'UV01' : 0.01, 'UV02' : 0.02, 'UV03' : 0.03, 'UV04' : 0.04 , 'UV05' : 0.05, 'UV06' : 0.06, 'UV07' : 0.07, 
                 'UV08' : 0.08, 'UV09' : 0.09, 'UV10' : 0.10, 'UV11' : 0.11, 'UV12' : 0.12, 'UV13' : 0.13, 'UV14' : 0.14,'UV15' : 0.15,'UVG04' : 0.04,'UVG03' : 0.03};             
    var authors = ['KA','TR','OG','AK'];
    var arr = [];
    var sheet = SpreadsheetApp.getActiveSheet();
    var row = 2;
    var campName;
    
    
    
    var rangeData = sheet.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    var lastRow = rangeData.getLastRow();

    let adNames = getColumnValuesArray('Ad Name');
    let titles = getColumnValuesArray('Title');
    let imageHashes = getColumnValuesArray('Image Hash');
    
    for (var i = 1; i < lastColumn + 1; i++) {
         arr.push(sheet.getRange(1, i).getValue());
    }
    
    let cmp = sheet.getRange(2, arr.indexOf('Campaign Name')+1).getValue();
    let cmpId = sheet.getRange(2,arr.indexOf('Campaign ID') + 1).getValue();
    let specialAdCategory = sheet.getRange(2, arr.indexOf('Special Ad Category') + 1).getValue();
    let campaignObjective = sheet.getRange(2, arr.indexOf('Campaign Objective') + 1).getValue();
    let buyingType = sheet.getRange(2, arr.indexOf('Buying Type') + 1).getValue();
    let adsetTimeStart = sheet.getRange(2, arr.indexOf('Ad Set Time Start') + 1).getDisplayValue();
    let destinationType = sheet.getRange(2, arr.indexOf('Destination Type') + 1).getValue();
    let adsetLifetimeBudget = sheet.getRange(2, arr.indexOf('Ad Set Lifetime Budget') + 1).getValue();
    let optimizedCTP = sheet.getRange(2, arr.indexOf('Optimized Conversion Tracking Pixels') + 1).getValue();
    let locationTypes = sheet.getRange(2, arr.indexOf('Location Types') + 1).getValue();
    let ageMin = sheet.getRange(2, arr.indexOf('Age Min') + 1).getValue();
    let ageMax = sheet.getRange(2, arr.indexOf('Age Max') + 1).getValue();
    let adStatus = sheet.getRange(2, arr.indexOf('Ad Status') + 1).getValue();
    //let instaPreview = sheet.getRange(2, arr.indexOf('Instagram Preview Link') + 1).getValue();
    let body = sheet.getRange(2, arr.indexOf('Body') + 1).getValue();
    let linkDescription = sheet.getRange(2, arr.indexOf('Link Description') + 1).getValue();
    let retailerIds = sheet.getRange(2, arr.indexOf('Retailer IDs') + 1).getValue();
    let posClickHeadline = sheet.getRange(2, arr.indexOf('Post Click Item Headline') + 1).getValue();
    let posClickDesc = sheet.getRange(2, arr.indexOf('Post Click Item Description') + 1).getValue();
    let displayLink = sheet.getRange(2, arr.indexOf('Display Link') + 1).getValue();
    let convPixels = sheet.getRange(2, arr.indexOf('Conversion Tracking Pixels') + 1).getValue();
    let imageCrops = sheet.getRange(2, arr.indexOf('Image Crops') + 1).getValue();
    let videoThumbnail = sheet.getRange(2, arr.indexOf('Video Thumbnail URL') + 1).getValue();
    let creativeType = sheet.getRange(2, arr.indexOf('Creative Type') + 1).getValue();
    let additionalCTS = sheet.getRange(2, arr.indexOf('Additional Custom Tracking Specs') + 1).getValue();
    let videoRetargeting = sheet.getRange(2, arr.indexOf('Video Retargeting') + 1).getValue();
    let usePageAsActor = sheet.getRange(2, arr.indexOf('Use Page as Actor') + 1).getValue();
    let useAD = sheet.getRange(2, arr.indexOf('Use Accelerated Delivery') + 1).getValue();
    let brandSafety = sheet.getRange(2, arr.indexOf('Brand Safety Inventory Filtering Levels') + 1).getValue();
    let adSetLifetimeImpressions = sheet.getRange(2, arr.indexOf('Ad Set Lifetime Impressions') + 1).getValue();

    
  for(let version of versions) {

    //  if(site == 'JR'){
    //    platform = 'Mobile';
    //  }

    let versionStringArray = version.split(' ');
    let versionName = versionStringArray[versionStringArray.length - 1];

    let counter = 0;
    
    for (let i = startPoint; i <= lastRow; i++) {

      let geo;
      let platform;
      let uv;
      let account;
      let author = 'none';
      let site;
      let name = [];
      let utm = '';

      for(let j = 0 ; j < sites.length ; j++){
        if(versionStringArray.indexOf(sites[j]) > -1){
          site = sites[j];
        }
      }
      
      for(let j = 0 ; j < uvs.length ; j++){
        if(versionStringArray.indexOf(uvs[j]) > -1){
          uv = uvs[j];
        }
      }

      for(let j = 0 ; j < geos.length ; j++){
        if(versionStringArray.indexOf(geos[j]) > -1){
          geo = geos[j];
        }
      }

      for(let j = 0 ; j < platforms.length ; j++){
        if(versionStringArray.indexOf(platforms[j]) > -1){
          platform = platforms[j];
        }
      }

      for(let j = 0 ; j < accounts.length ; j++){
        if(versionStringArray.indexOf(accounts[j]) > -1){
          account = accounts[j];
        }
      }

      for(let j = 0 ; j < authors.length ; j++){
        if(versionStringArray.indexOf(authors[j]) > -1) {
          author = authors[j];
        }
      }

      for(let j = 0 ; j < uvs.length ; j++){
        if(versionStringArray.indexOf(uvs[j]) > -1) {
          uv = uvs[j];
        }
      }

      

      if(author != 'none'){
        for(let j = versionStringArray.indexOf('-') + 1 ; j < versionStringArray.indexOf(author) ; j++){
        name.push(versionStringArray[j]);
        }
      }else{
        for(let j = versionStringArray.indexOf('-') + 1 ; j < versionStringArray.indexOf(platform) ; j++){
        name.push(versionStringArray[j]);
      }
      }

      name = name.join('+');
    
      let ver = versionStringArray[versionStringArray.indexOf(geo) + 1];

      
      if(author != 'none') {
        utm = name + '+' + author + '+' + geo + '+' + platform + '+' + ver + '+' + uv + '+' + account;
      } else{
          utm = name + '+' + geo + '+' + platform + '+' + ver + '+' + uv + '+' + account;
      }

      if(utm.length >= 40){
        let ui = SpreadsheetApp.getUi();
        ui.alert('UTM too long. Try again.',ui.ButtonSet.OK);
        return;
      }

      utm = 'utm_source=Facebook&utm_medium=Facebook&utm_campaign=' + utm;
    
      if(site == 'JR' && account != '15' && account != 'III' && account != '20'){
        let ui = SpreadsheetApp.getUi();
        ui.alert('JourneyRanger campaigns should be only uploaded to accounts: III,15,20.',ui.ButtonSet.OK);
        return;
      }
    
      if(site == 'TF' && (account == '15' || account == 'III' || account == '20')){
        let ui = SpreadsheetApp.getUi();
        ui.alert('TeddyFeed campaigns should not be uploaded to accounts: III,15,20.',ui.ButtonSet.OK);
        return;
      }
    
      let link = sheet.getRange(2, arr.indexOf('Link')+1).getValue().split('?');
      link[link.length-1] = utm;
      let str = link.join('?');
      
      sheet.getRange(i, arr.indexOf('Campaign Status')+1).setValue('ACTIVE');
      sheet.getRange(i, arr.indexOf('Ad Set Run Status')+1).setValue('ACTIVE');
      //sheet.getRange(i, arr.indexOf('Campaign ID')+1).setValue('');
      sheet.getRange(i, arr.indexOf('Ad Set ID')+1).setValue('');
      sheet.getRange(i, arr.indexOf('Ad ID')+1).setValue('');
      sheet.getRange(i, arr.indexOf('Campaign ID') + 1).setValue(cmpId);
      sheet.getRange(i, arr.indexOf('Campaign Name')+1).setValue(cmp);
      sheet.getRange(i, arr.indexOf('Special Ad Category') + 1).setValue(specialAdCategory);
      sheet.getRange(i, arr.indexOf('Campaign Objective') + 1).setValue(campaignObjective);
      sheet.getRange(i, arr.indexOf('Buying Type') + 1).setValue(buyingType);
      sheet.getRange(i, arr.indexOf('Ad Set Time Start') + 1).setValue(adsetTimeStart);
      sheet.getRange(i, arr.indexOf('Destination Type') + 1).setValue(destinationType);
      sheet.getRange(i, arr.indexOf('Ad Set Lifetime Budget') + 1).setValue(adsetLifetimeBudget);
      sheet.getRange(i, arr.indexOf('Optimized Conversion Tracking Pixels') + 1).setValue(optimizedCTP);
      sheet.getRange(i, arr.indexOf('Location Types') + 1).setValue(locationTypes);
      sheet.getRange(i, arr.indexOf('Age Min') + 1).setValue(ageMin);
      sheet.getRange(i, arr.indexOf('Age Max') + 1).setValue(ageMax);
      sheet.getRange(i, arr.indexOf('Ad Status') + 1).setValue(adStatus);
      //sheet.getRange(i, arr.indexOf('Instagram Preview Link') + 1).setValue(instaPreview);
      
      sheet.getRange(i, arr.indexOf('Ad Name') + 1).setValue(adNames[counter]);
      sheet.getRange(i, arr.indexOf('Title') + 1).setValue(titles[counter]);
      sheet.getRange(i, arr.indexOf('Image Hash') + 1).setValue(imageHashes[counter]);
      counter++;
      
      sheet.getRange(i, arr.indexOf('Body') + 1).setValue(body);
      sheet.getRange(i, arr.indexOf('Link Description') + 1).setValue(linkDescription);
      sheet.getRange(i, arr.indexOf('Retailer IDs') + 1).setValue(retailerIds);
      sheet.getRange(i, arr.indexOf('Post Click Item Headline') + 1).setValue(posClickHeadline);
      sheet.getRange(i, arr.indexOf('Post Click Item Description') + 1).setValue(posClickDesc);
      sheet.getRange(i, arr.indexOf('Display Link') + 1).setValue(displayLink);
      sheet.getRange(i, arr.indexOf('Conversion Tracking Pixels') + 1).setValue(convPixels);
      sheet.getRange(i, arr.indexOf('Image Crops') + 1).setValue(imageCrops);
      sheet.getRange(i, arr.indexOf('Video Thumbnail URL') + 1).setValue(videoThumbnail);
      sheet.getRange(i, arr.indexOf('Creative Type') + 1).setValue(creativeType);
      sheet.getRange(i, arr.indexOf('Additional Custom Tracking Specs') + 1).setValue(additionalCTS);
      sheet.getRange(i, arr.indexOf('Video Retargeting') + 1).setValue(videoRetargeting);
      sheet.getRange(i, arr.indexOf('Use Page as Actor') + 1).setValue(usePageAsActor);
      sheet.getRange(i, arr.indexOf('Use Accelerated Delivery') + 1).setValue(useAD);
      sheet.getRange(i, arr.indexOf('Brand Safety Inventory Filtering Levels') + 1).setValue(brandSafety);
      sheet.getRange(i, arr.indexOf('Ad Set Name')+1).setValue(version);
      sheet.getRange(i, arr.indexOf('Ad Set Daily Budget')+1).setValue(50);
      sheet.getRange(i, arr.indexOf('Optimized Pixel Rule')+1).setValue('{"and":[{"event":{"eq":"UV"}},{"or":[{"value":{"eq":"' + uvMap[uv] + '"}}]}]}');
      sheet.getRange(i, arr.indexOf('Optimized Event')+1).setValue('OTHER');
      sheet.getRange(i, arr.indexOf('Link')+1).setValue(`${str}+${versionName}+FB`);
      sheet.getRange(i, arr.indexOf('Optimization Goal')+1).setValue('OFFSITE_CONVERSIONS');
      sheet.getRange(i, arr.indexOf('Attribution Spec')+1).setValue('[{"event_type":"CLICK_THROUGH","window_days":1},{"event_type":"VIEW_THROUGH","window_days":1}]');
      sheet.getRange(i, arr.indexOf('Billing Event')+1).setValue('IMPRESSIONS');
      sheet.getRange(i, arr.indexOf('Bid Amount')+1).setValue('');
      sheet.getRange(i, arr.indexOf('Ad Set Bid Strategy')+1).setValue('Lowest Cost');
      sheet.getRange(i, arr.indexOf('Ad Set Lifetime Impressions') + 1).setValue(adSetLifetimeImpressions)
      
      if(site == 'TF'){
        sheet.getRange(i, arr.indexOf('Instagram Account ID')+1).setValue('');//x:1677001952358894
        sheet.getRange(i, arr.indexOf('Link Object ID')+1).setValue('o:115777153153104');}
      if(site == 'JR'){
        sheet.getRange(i, arr.indexOf('Instagram Account ID')+1).setValue('x:1993512290745708');
        sheet.getRange(i, arr.indexOf('Link Object ID')+1).setValue('o:286148318924709');}
      if(site == 'BB'){
        sheet.getRange(i, arr.indexOf('Instagram Account ID')+1).setValue('x:1953113638140064');
        sheet.getRange(i, arr.indexOf('Link Object ID')+1).setValue('o:274748396529829');}
      
      if(arr.indexOf('Minimum ROAS') > -1){sheet.getRange(i, arr.indexOf('Minimum ROAS')+1).setValue('');}
      
      if(geo == 'WW'){
        sheet.getRange(i, arr.indexOf('Global Regions')+1).setValue('worldwide'); 
        sheet.getRange(i, arr.indexOf('Excluded Global Regions')+1).setValue('europe');}
      if(geo == 'CAUS' || geo == 'CAAUUS'){
        sheet.getRange(i, arr.indexOf('Countries')+1).setValue('US, CA, AU');
        sheet.getRange(i, arr.indexOf('Locales')+1).setValue('');}
      if(geo == 'US'){
        sheet.getRange(i, arr.indexOf('Countries')+1).setValue('US');
        sheet.getRange(i, arr.indexOf('Locales')+1).setValue('');}
      if(geo == 'CA'){
        sheet.getRange(i, arr.indexOf('Countries')+1).setValue('CA');
        sheet.getRange(i, arr.indexOf('Locales')+1).setValue('English (UK), English (US)');}
      if(geo == 'AU'){
        sheet.getRange(i, arr.indexOf('Countries')+1).setValue('AU');
        sheet.getRange(i, arr.indexOf('Locales')+1).setValue('');}
      if(geo == 'NAC'){
        sheet.getRange(i, arr.indexOf('Countries')+1).setValue('CA, AU, NZ');
        sheet.getRange(i, arr.indexOf('Locales')+1).setValue('English (UK), English (US)');}
      if(geo == 'CAU'){
        sheet.getRange(i, arr.indexOf('Countries')+1).setValue('CA, AU');
        sheet.getRange(i, arr.indexOf('Locales')+1).setValue('English (UK), English (US)');}  
      
      if(platform == 'Desktop'){
        sheet.getRange(i, arr.indexOf('Publisher Platforms')+1).setValue('facebook, audience_network'); 
        sheet.getRange(i, arr.indexOf('Facebook Positions')+1).setValue('feed, right_hand_column, instream_video, marketplace'); 
        sheet.getRange(i, arr.indexOf('Instagram Positions')+1).setValue(''); 
        sheet.getRange(i, arr.indexOf('Audience Network Positions')+1).setValue('instream_video');
        sheet.getRange(i, arr.indexOf('Messenger Positions')+1).setValue('');
        sheet.getRange(i, arr.indexOf('Device Platforms')+1).setValue('desktop');
        sheet.getRange(i, arr.indexOf('User Operating System')+1).setValue('');}
                                
      if(platform == 'MobA'){
        sheet.getRange(i, arr.indexOf('Publisher Platforms')+1).setValue('facebook, instagram, audience_network, messenger'); 
        sheet.getRange(i, arr.indexOf('Facebook Positions')+1).setValue('feed, right_hand_column, video_feeds, instant_article, instream_video, marketplace, story, search'); 
        sheet.getRange(i, arr.indexOf('Instagram Positions')+1).setValue('stream, story, explore'); 
        sheet.getRange(i, arr.indexOf('Audience Network Positions')+1).setValue('classic, instream_video, rewarded_video');
        sheet.getRange(i, arr.indexOf('Messenger Positions')+1).setValue('messenger_home, story');
        sheet.getRange(i, arr.indexOf('Device Platforms')+1).setValue('mobile');
        sheet.getRange(i, arr.indexOf('User Operating System')+1).setValue('');}
                                
      if(platform == 'Auto' || platform == 'AutoAnd'){
        sheet.getRange(i, arr.indexOf('Publisher Platforms')+1).setValue(''); 
        sheet.getRange(i, arr.indexOf('Facebook Positions')+1).setValue(''); 
        sheet.getRange(i, arr.indexOf('Instagram Positions')+1).setValue(''); 
        sheet.getRange(i, arr.indexOf('Audience Network Positions')+1).setValue('');
        sheet.getRange(i, arr.indexOf('Messenger Positions')+1).setValue('');
        sheet.getRange(i, arr.indexOf('Device Platforms')+1).setValue('');
        sheet.getRange(i, arr.indexOf('User Operating System')+1).setValue('');
      }
    }

    startPoint = SpreadsheetApp.getActiveSheet().getDataRange().getLastRow() + 1;
    lastRow += rangeData.getLastRow() - 1;
    
  }
  
  SpreadsheetApp.getUi().alert('Operation Complete!');
}

function getColumnValuesArray(columnName) {
  let columnValuesArray = [];
  let numberOfColumns = SpreadsheetApp.getActiveSheet().getDataRange().getLastColumn();
  let numberOfRows = SpreadsheetApp.getActiveSheet().getDataRange().getLastRow();

  for(let columnIndex = 1; columnIndex <= numberOfColumns; columnIndex++) {
    let currentColumnName = SpreadsheetApp.getActiveSheet().getRange(1, columnIndex).getValue();

    if(columnName === currentColumnName) {
      let columnValuesTwoDimensionalArray = SpreadsheetApp.getActiveSheet().getRange(2, columnIndex, numberOfRows - 1).getValues();
      columnValuesTwoDimensionalArray.forEach(value => columnValuesArray.push(value[0]));

      return columnValuesArray;
    }
  }

  return columnValuesArray;
}

function campDup(){
  //var geos = ['US', 'CA', 'AU', 'WW', 'CAUS' , 'CAAUUS','NAC','CAU'];
  var platforms = ['Mob', 'Desk', 'Auto','AutoAnd','MobA','IOS'];
  var sites = ['TF','JR','BB'];
  var accounts = ['I', 'II', 'III', 'IV', 'V', 'TF1', 'TF2', 'TF3', '6', '10', '11', '12', '14', '15', '16', '17', '18', '19', '20'];
  var uvs = ['UV01', 'UV02', 'UV03', 'UV04', 'UV05', 'UV06', 'UV07', 'UV08', 'UV09', 'UV10', 'UV11', 'UV12', 'UV13', 'UV14', 'UV15','UVG04','UVG03','UV04F'];
  var uvMap = {'UV01' : 0.01, 'UV02' : 0.02, 'UV03' : 0.03, 'UV04' : 0.04 , 'UV05' : 0.05, 'UV06' : 0.06, 'UV07' : 0.07, 
               'UV08' : 0.08, 'UV09' : 0.09, 'UV10' : 0.10, 'UV11' : 0.11, 'UV12' : 0.12, 'UV13' : 0.13, 'UV14' : 0.14,'UV15' : 0.15,'UVG04' : 0.04,'UVG03' : 0.03};             
  var authors = ['KA','TR','OG','AK'];
  var arr = [];
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = 2;
  var campName;
  var name = [];
  var geo;
  var platform;
  var ver;
  var uv;
  var account;
  var author = 'none';
  var utm;
  var site;
  
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  
  for (var i = 1; i < lastColumn + 1; i++) {
       arr.push(sheet.getRange(1, i).getValue());
  }
  
  var cmp = sheet.getRange(2, arr.indexOf('Campaign Name')+1).getValue();
  campName = (sheet.getRange(2, arr.indexOf('Campaign Name')+1).getValue()).split(' ');
  
  for(var i = 0 ; i < sites.length ; i++){
    if(campName.indexOf(sites[i]) > -1){
      site = sites[i];
    }
  }
  
  for(var i = 0 ; i < geos.length ; i++){
    if(campName.indexOf(geos[i]) > -1){
      geo = geos[i];
    }
  }
  
  for(var i = 0 ; i < platforms.length ; i++){
    if(campName.indexOf(platforms[i]) > -1){
      platform = platforms[i];
    }
  }
  
  
  for(var i = 0 ; i < accounts.length ; i++){
    if(campName.indexOf(accounts[i]) > -1){
      account = accounts[i];
    }
  }
  
  for(var i = 0 ; i < authors.length ; i++){
    if(campName.indexOf(authors[i]) > -1){
      author = authors[i];
    }
  }
  
  for(var i = 0 ; i < uvs.length ; i++){
    if(campName.indexOf(uvs[i]) > -1){
      uv = uvs[i];
    }
  }
  
  if(author != 'none'){
    for(var i = campName.indexOf('-')+1 ; i < campName.indexOf(author) ; i++){
    name.push(campName[i]);
    }
  }else{
    for(var i = campName.indexOf('-')+1 ; i < campName.indexOf(platform) ; i++){
    name.push(campName[i]);
  }
  }
  
  
  name = name.join('+');
  
  ver = campName[campName.indexOf(geo)+1];
  
  if(author != 'none'){
  utm = name + '+' + author + '+' + geo + '+' + platform + '+' + ver + '+' + uv + '+' + account + '+FB';
  }else{
    utm = name + '+' + geo + '+' + platform + '+' + ver + '+' + uv + '+' + account + '+FB';
  }
  
  if(utm.length >= 40){
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('UTM too long. Try again.',ui.ButtonSet.OK);
    return;
  }
  
  utm = 'utm_source=Facebook&utm_medium=Facebook&utm_campaign='+utm;
  
  if(site == 'JR' && account != '15' && account != 'III' && account != '20'){
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('JourneyRanger campaigns should be only uploaded to accounts: III,15,20.',ui.ButtonSet.OK);
    return;
  }
  
  if(site == 'TF' && (account == '15' || account == 'III' || account == '20')){
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('TeddyFeed campaigns should not be uploaded to accounts: III,15,20.',ui.ButtonSet.OK);
    return;
  }
  
  var link = sheet.getRange(2, arr.indexOf('Link')+1).getValue().split('?');
  link[link.length-1] = utm;
  var str = link.join('?');
  
//  if(site == 'JR'){
//    platform = 'Mobile';
//  }
  
  for (var i = 2; i < lastRow + 1; i++) {
    sheet.getRange(i, arr.indexOf('Campaign Status')+1).setValue('ACTIVE');
    sheet.getRange(i, arr.indexOf('Ad Set Run Status')+1).setValue('ACTIVE');
    //sheet.getRange(i, arr.indexOf('Campaign ID')+1).setValue('');
    sheet.getRange(i, arr.indexOf('Ad Set ID')+1).setValue('');
    sheet.getRange(i, arr.indexOf('Ad ID')+1).setValue('');
    sheet.getRange(i, arr.indexOf('Campaign Name')+1).setValue(cmp);
    sheet.getRange(i, arr.indexOf('Ad Set Name')+1).setValue(cmp);
    sheet.getRange(i, arr.indexOf('Ad Set Daily Budget')+1).setValue(50);
    sheet.getRange(i, arr.indexOf('Optimized Pixel Rule')+1).setValue('{"and":[{"event":{"eq":"UV"}},{"or":[{"value":{"eq":"' + uvMap[uv] + '"}}]}]}');
    sheet.getRange(i, arr.indexOf('Optimized Event')+1).setValue('OTHER');
    sheet.getRange(i, arr.indexOf('Link')+1).setValue(str);
    sheet.getRange(i, arr.indexOf('Optimization Goal')+1).setValue('OFFSITE_CONVERSIONS');
    sheet.getRange(i, arr.indexOf('Attribution Spec')+1).setValue('[{"event_type":"CLICK_THROUGH","window_days":1},{"event_type":"VIEW_THROUGH","window_days":1}]');
    sheet.getRange(i, arr.indexOf('Billing Event')+1).setValue('IMPRESSIONS');
    sheet.getRange(i, arr.indexOf('Bid Amount')+1).setValue('');
    sheet.getRange(i, arr.indexOf('Ad Set Bid Strategy')+1).setValue('Lowest Cost');
    
    if(site == 'TF'){
      sheet.getRange(i, arr.indexOf('Instagram Account ID')+1).setValue('');//x:1677001952358894
      sheet.getRange(i, arr.indexOf('Link Object ID')+1).setValue('o:115777153153104');}
    if(site == 'JR'){
      sheet.getRange(i, arr.indexOf('Instagram Account ID')+1).setValue('x:1993512290745708');
      sheet.getRange(i, arr.indexOf('Link Object ID')+1).setValue('o:286148318924709');}
    if(site == 'BB'){
      sheet.getRange(i, arr.indexOf('Instagram Account ID')+1).setValue('x:1953113638140064');
      sheet.getRange(i, arr.indexOf('Link Object ID')+1).setValue('o:274748396529829');}
    
    if(arr.indexOf('Minimum ROAS') > -1){sheet.getRange(i, arr.indexOf('Minimum ROAS')+1).setValue('');}
    
    if(geo == 'WW'){
      sheet.getRange(i, arr.indexOf('Global Regions')+1).setValue('worldwide'); 
      sheet.getRange(i, arr.indexOf('Excluded Global Regions')+1).setValue('europe');}
    if(geo == 'CAUS' || geo === 'CAAUUS'){
      sheet.getRange(i, arr.indexOf('Countries')+1).setValue('US, CA, AU');
      sheet.getRange(i, arr.indexOf('Locales')+1).setValue('');}
    if(geo == 'US'){
      sheet.getRange(i, arr.indexOf('Countries')+1).setValue('US');
      sheet.getRange(i, arr.indexOf('Locales')+1).setValue('');}
    if(geo == 'CA'){
      sheet.getRange(i, arr.indexOf('Countries')+1).setValue('CA');
      sheet.getRange(i, arr.indexOf('Locales')+1).setValue('English (UK), English (US)');}
    if(geo == 'AU'){
      sheet.getRange(i, arr.indexOf('Countries')+1).setValue('AU');
      sheet.getRange(i, arr.indexOf('Locales')+1).setValue('');}
    if(geo == 'NAC'){
      sheet.getRange(i, arr.indexOf('Countries')+1).setValue('CA, AU, NZ');
      sheet.getRange(i, arr.indexOf('Locales')+1).setValue('English (UK), English (US)');}
    if(geo == 'CAU'){
      sheet.getRange(i, arr.indexOf('Countries')+1).setValue('CA, AU');
      sheet.getRange(i, arr.indexOf('Locales')+1).setValue('English (UK), English (US)');}  
    
    if(platform == 'Desktop'){
      sheet.getRange(i, arr.indexOf('Publisher Platforms')+1).setValue('facebook, audience_network'); 
      sheet.getRange(i, arr.indexOf('Facebook Positions')+1).setValue('feed, right_hand_column, instream_video, marketplace'); 
      sheet.getRange(i, arr.indexOf('Instagram Positions')+1).setValue(''); 
      sheet.getRange(i, arr.indexOf('Audience Network Positions')+1).setValue('instream_video');
      sheet.getRange(i, arr.indexOf('Messenger Positions')+1).setValue('');
      sheet.getRange(i, arr.indexOf('Device Platforms')+1).setValue('desktop');
      sheet.getRange(i, arr.indexOf('User Operating System')+1).setValue('');}
                              
    if(platform == 'MobA'){
      sheet.getRange(i, arr.indexOf('Publisher Platforms')+1).setValue('facebook, instagram, audience_network, messenger'); 
      sheet.getRange(i, arr.indexOf('Facebook Positions')+1).setValue('feed, right_hand_column, video_feeds, instant_article, instream_video, marketplace, story, search'); 
      sheet.getRange(i, arr.indexOf('Instagram Positions')+1).setValue('stream, story, explore'); 
      sheet.getRange(i, arr.indexOf('Audience Network Positions')+1).setValue('classic, instream_video, rewarded_video');
      sheet.getRange(i, arr.indexOf('Messenger Positions')+1).setValue('messenger_home, story');
      sheet.getRange(i, arr.indexOf('Device Platforms')+1).setValue('mobile');
      sheet.getRange(i, arr.indexOf('User Operating System')+1).setValue('');}
                             
    if(platform == 'Auto' || platform === 'AutoAnd'){
      sheet.getRange(i, arr.indexOf('Publisher Platforms')+1).setValue(''); 
      sheet.getRange(i, arr.indexOf('Facebook Positions')+1).setValue(''); 
      sheet.getRange(i, arr.indexOf('Instagram Positions')+1).setValue(''); 
      sheet.getRange(i, arr.indexOf('Audience Network Positions')+1).setValue('');
      sheet.getRange(i, arr.indexOf('Messenger Positions')+1).setValue('');
      sheet.getRange(i, arr.indexOf('Device Platforms')+1).setValue('');
      sheet.getRange(i, arr.indexOf('User Operating System')+1).setValue('');}
    
  }
  
}
