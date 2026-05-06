/**
 * ADAM スタッフ報酬管理 - Google Apps Script バックエンド
 *
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成
 * 2. スプレッドシートのIDをコピーし、下の SPREADSHEET_ID に貼り付け
 * 3. 拡張機能 → Apps Script を開く
 * 4. このコードを貼り付けて保存
 * 5. デプロイ → 新しいデプロイ → ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員
 * 6. デプロイ後に表示されるURLをアプリに入力
 * 7. （任意）setupWeeklyTrigger を1回手動実行 → 週次バックアップを有効化
 */

function doGet(e){return handleRequest(e)}
function doPost(e){return handleRequest(e)}

function handleRequest(e){
  const action=(e&&e.parameter&&e.parameter.action)||'';
  // 読み取りはロック不要（同時アクセスでブロックされないように）
  if(action==='getData'){
    try{return jsonResponse(getAllData())}
    catch(err){return jsonResponse({error:err.message})}
  }
  // 書き込みのみロック
  const lock=LockService.getScriptLock();
  try{lock.waitLock(15000)}catch(err){
    return jsonResponse({error:'BUSY'});
  }
  try{
    if(action==='saveData'){
      const body=JSON.parse(e.postData.contents);
      saveAllData(body);
      return jsonResponse({success:true});
    }
    return jsonResponse({error:'Unknown action: '+action});
  }catch(err){
    return jsonResponse({error:err.message});
  }finally{
    lock.releaseLock();
  }
}

function jsonResponse(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===== Sheet Helpers ===== */
const SPREADSHEET_ID='1U5IRUVWp_GMjIFoejoVmi6EXfgZRFqtaRb9-zcXicF8';
function getSpreadsheet(){
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function getOrCreateSheet(name,headers){
  const ss=getSpreadsheet();
  let sh=ss.getSheetByName(name);
  if(!sh){
    sh=ss.insertSheet(name);
    if(headers&&headers.length>0){
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    }
  }else if(headers&&headers.length>0){
    // ヘッダー行が短ければ不足分を補う（後方互換のための列追加）
    const lastCol=Math.max(sh.getLastColumn(),1);
    const existing=sh.getRange(1,1,1,lastCol).getValues()[0];
    if(existing.length<headers.length){
      sh.getRange(1,existing.length+1,1,headers.length-existing.length).setValues([headers.slice(existing.length)]);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    }
  }
  return sh;
}

/** スプレッドシートが日付に自動変換した月キーをYYYY-MM形式に戻す */
function normalizeMonthKey(val){
  var s=String(val);
  if(/^\d{4}-\d{2}$/.test(s))return s;
  var d;
  if(val instanceof Date){d=val}
  else{d=new Date(s); if(isNaN(d.getTime()))return s}
  var y=d.getFullYear();
  var m=String(d.getMonth()+1).padStart(2,'0');
  return y+'-'+m;
}

function sheetToArray(sh){
  const data=sh.getDataRange().getValues();
  if(data.length<=1)return[];
  const headers=data[0];
  const rows=[];
  for(let i=1;i<data.length;i++){
    const obj={};
    headers.forEach((h,j)=>{obj[h]=data[i][j]});
    rows.push(obj);
  }
  return rows;
}

/* ===== Get All Data ===== */
function getAllData(){
  // Config
  const cfgSh=getOrCreateSheet('設定',['key','value']);
  const cfgRows=sheetToArray(cfgSh);
  const config={
    masterPin:'0000',
    staff:[],
    prices:{
      aden:{room:8000},
      'sub-svc':{room:5000},
      noic:{ad:6000,self:10000,monitor:2500},
      resync:{ad:6000,self:10000,monitor:2500}
    }
  };
  // 旧 trainerPin/trainerName を staff[0] に移行するための一時保持
  let legacyTrainerPin=null,legacyTrainerName=null;
  cfgRows.forEach(r=>{
    if(r.key==='masterPin')config.masterPin=String(r.value).padStart(4,'0');
    else if(r.key==='trainerPin')legacyTrainerPin=String(r.value).padStart(4,'0');
    else if(r.key==='trainerName')legacyTrainerName=String(r.value);
    else if(r.key==='staff')try{config.staff=JSON.parse(r.value)||[]}catch(e){}
    else if(r.key==='prices')try{config.prices=JSON.parse(r.value)}catch(e){}
  });
  if((legacyTrainerName||legacyTrainerPin)&&config.staff.length===0){
    config.staff.push({
      id:'s_'+Date.now()+'_legacy',
      name:legacyTrainerName||'スタッフ1',
      pin:legacyTrainerPin||'0000',
      createdAt:Date.now()
    });
  }

  // Clients
  const clSh=getOrCreateSheet('ゲスト',['id','name','service','source','sessionPrice','courses','createdAt','staffId','customEnabled','customSessionPrice','customReward']);
  const defaultStaffId=config.staff[0]?config.staff[0].id:null;
  const clients=sheetToArray(clSh).map(r=>({
    id:String(r.id),
    name:String(r.name),
    service:String(r.service),
    source:r.source?String(r.source):null,
    sessionPrice:Number(r.sessionPrice)||0,
    courses:r.courses?JSON.parse(r.courses):[],
    createdAt:Number(r.createdAt)||0,
    staffId:r.staffId?String(r.staffId):defaultStaffId,
    customEnabled:r.customEnabled===true||r.customEnabled==='true'||r.customEnabled===1,
    customSessionPrice:Number(r.customSessionPrice)||0,
    customReward:Number(r.customReward)||0
  }));

  // Sessions
  const seSh=getOrCreateSheet('セッション',['month','clientId','count','sessionPriceOverride']);
  const seRows=sheetToArray(seSh);
  const sessions={};
  seRows.forEach(r=>{
    const mk=normalizeMonthKey(r.month);
    const cid=String(r.clientId);
    if(!sessions[mk])sessions[mk]={};
    sessions[mk][cid]={
      count:Number(r.count)||0,
      sessionPriceOverride:r.sessionPriceOverride!==''&&r.sessionPriceOverride!=null?Number(r.sessionPriceOverride):null
    };
  });

  // Locked Months
  const lkSh=getOrCreateSheet('確定月',['month','lockedAt']);
  const lkRows=sheetToArray(lkSh);
  const lockedMonths={};
  lkRows.forEach(r=>{lockedMonths[normalizeMonthKey(r.month)]={lockedAt:Number(r.lockedAt)||0}});

  // Adjustments（報酬調整）
  const adjSh=getOrCreateSheet('報酬調整',['month','staffKey','enabled','amount','reason']);
  const adjRows=sheetToArray(adjSh);
  const adjustments={};
  adjRows.forEach(r=>{
    const mk=normalizeMonthKey(r.month);
    const sk=String(r.staffKey);
    if(!adjustments[mk])adjustments[mk]={};
    adjustments[mk][sk]={
      enabled:r.enabled===true||r.enabled==='true'||r.enabled===1,
      amount:Number(r.amount)||0,
      reason:String(r.reason||'')
    };
  });

  return{config,clients,sessions,lockedMonths,adjustments};
}

/* ===== Save All Data ===== */
function saveAllData(D){
  // Config
  const cfgSh=getOrCreateSheet('設定',['key','value']);
  cfgSh.getRange(2,1,Math.max(cfgSh.getLastRow(),2),2).clearContent();
  cfgSh.getRange(2,2,10,1).setNumberFormat('@');
  const cfgData=[
    ['masterPin',String(D.config.masterPin).padStart(4,'0')],
    ['staff',JSON.stringify(D.config.staff||[])],
    ['prices',JSON.stringify(D.config.prices)]
  ];
  cfgSh.getRange(2,1,cfgData.length,2).setValues(cfgData);

  // Clients
  const clSh=getOrCreateSheet('ゲスト',['id','name','service','source','sessionPrice','courses','createdAt','staffId','customEnabled','customSessionPrice','customReward']);
  if(clSh.getLastRow()>1)clSh.getRange(2,1,clSh.getLastRow()-1,11).clearContent();
  if(D.clients&&D.clients.length>0){
    const clData=D.clients.map(c=>[
      c.id,c.name,c.service,c.source||'',c.sessionPrice||0,
      JSON.stringify(c.courses||[]),c.createdAt||0,c.staffId||'',
      c.customEnabled?true:false,c.customSessionPrice||0,c.customReward||0
    ]);
    clSh.getRange(2,1,clData.length,11).setValues(clData);
  }

  // Sessions
  const seSh=getOrCreateSheet('セッション',['month','clientId','count','sessionPriceOverride']);
  if(seSh.getLastRow()>1)seSh.getRange(2,1,seSh.getLastRow()-1,4).clearContent();
  const seData=[];
  if(D.sessions){
    Object.keys(D.sessions).forEach(mk=>{
      Object.keys(D.sessions[mk]).forEach(cid=>{
        const s=D.sessions[mk][cid];
        seData.push([mk,cid,s.count||0,s.sessionPriceOverride!=null?s.sessionPriceOverride:'']);
      });
    });
  }
  if(seData.length>0){
    seSh.getRange(2,1,seData.length,1).setNumberFormat('@');
    seSh.getRange(2,1,seData.length,4).setValues(seData);
  }

  // Locked Months
  const lkSh=getOrCreateSheet('確定月',['month','lockedAt']);
  if(lkSh.getLastRow()>1)lkSh.getRange(2,1,lkSh.getLastRow()-1,2).clearContent();
  const lkData=[];
  if(D.lockedMonths){
    Object.keys(D.lockedMonths).forEach(mk=>{
      lkData.push([mk,D.lockedMonths[mk].lockedAt||0]);
    });
  }
  if(lkData.length>0){
    lkSh.getRange(2,1,lkData.length,1).setNumberFormat('@');
    lkSh.getRange(2,1,lkData.length,2).setValues(lkData);
  }

  // Adjustments
  const adjSh=getOrCreateSheet('報酬調整',['month','staffKey','enabled','amount','reason']);
  if(adjSh.getLastRow()>1)adjSh.getRange(2,1,adjSh.getLastRow()-1,5).clearContent();
  const adjData=[];
  if(D.adjustments){
    Object.keys(D.adjustments).forEach(mk=>{
      Object.keys(D.adjustments[mk]).forEach(sk=>{
        const a=D.adjustments[mk][sk];
        adjData.push([mk,sk,a.enabled?true:false,a.amount||0,a.reason||'']);
      });
    });
  }
  if(adjData.length>0){
    adjSh.getRange(2,1,adjData.length,1).setNumberFormat('@');
    adjSh.getRange(2,1,adjData.length,5).setValues(adjData);
  }
}

/* ===== Weekly Backup to Google Drive ===== */
const BACKUP_FOLDER_NAME='adam-staff-salary_バックアップ';
const BACKUP_KEEP_DAYS=90;

function getOrCreateBackupFolder(){
  const folders=DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if(folders.hasNext())return folders.next();
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

/**
 * 毎週自動実行：スプレッドシートの全データをJSONでGoogle Driveに保存
 */
function weeklyBackup(){
  try{
    const data=getAllData();
    const now=new Date();
    const dateStr=Utilities.formatDate(now,'Asia/Tokyo','yyyy-MM-dd_HHmm');
    const fileName='adam-staff-salary_'+dateStr+'.json';
    const json=JSON.stringify(data,null,2);
    const folder=getOrCreateBackupFolder();
    folder.createFile(fileName,json,'application/json');
    cleanOldBackups();
    Logger.log('バックアップ完了: '+fileName);
  }catch(e){
    Logger.log('バックアップ失敗: '+e.message);
    try{
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        '【adam-staff-salary】バックアップ失敗',
        '日時: '+new Date().toLocaleString()+'\nエラー: '+e.message);
    }catch(me){}
  }
}

function cleanOldBackups(){
  const folder=getOrCreateBackupFolder();
  const cutoff=new Date();
  cutoff.setDate(cutoff.getDate()-BACKUP_KEEP_DAYS);
  const files=folder.getFiles();
  while(files.hasNext()){
    const f=files.next();
    if(f.getDateCreated()<cutoff){
      f.setTrashed(true);
      Logger.log('古いバックアップを削除: '+f.getName());
    }
  }
}

/**
 * 初回1回だけ手動実行：毎週月曜9時のトリガーを設定
 */
function setupWeeklyTrigger(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()==='weeklyBackup')ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyBackup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
  Logger.log('週次バックアップトリガーを設定しました（毎週月曜 9:00〜10:00）');
}
