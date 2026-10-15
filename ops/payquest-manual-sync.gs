/**
 * PAY QUEST: 手動一方向同期 (Supabase -> Google Sheets)
 * Google Apps Script にこのファイルを貼り、スクリプトプロパティ
 * SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY を設定して syncPayQuestManual を実行。
 * このファイル・GitHub・公開Webページに秘密鍵を記載しない。
 * トリガーを作成しない。Supabaseへの書き込みを行わない。
 */
const PQ_SHEET_ID = '1zruRxrIe3iqgCshGKp7T8jNFmWvAv2F3IBUpLOvr7lc';
const PQ_STATUS = {
 SUBMITTED:'見積中',REVIEWING:'見積中',QUOTE_READY:'見積中',WAITING_APPROVAL:'見積中',REQUOTING:'見積中',
 SECURITY_CHECK:'決済中',PAYMENT_PENDING:'決済中',PARTIALLY_PAID:'決済中',
 QUEST_CLEAR:'調達中',PURCHASING:'調達中',DELIVERY:'発送済',
 COMPLETE:'COMPLETE',CANCELLED:'CANCELLED',REFUND_PENDING:'CANCELLED',REFUNDED:'CANCELLED'
};
function pqFetch_(table, cols) {
 const props=PropertiesService.getScriptProperties();
 const url=props.getProperty('SUPABASE_URL');
 const key=props.getProperty('SUPABASE_SERVICE_ROLE_KEY');
 if(!url || !key || !/^https:\/\/[a-z0-9-]+\.supabase\.co$/.test(url)) throw Error('Supabase接続設定が不足または不正');
 let out=[],offset=0;
 while(true){
  const endpoint=url+'/rest/v1/'+table+'?select='+encodeURIComponent(cols)+'&order=created_at.asc,id.asc&limit=1000&offset='+offset;
  const r=UrlFetchApp.fetch(endpoint,{headers:{apikey:key,Authorization:'Bearer '+key},muteHttpExceptions:true});
  if(r.getResponseCode()!==200) throw Error(table+'取得失敗 HTTP '+r.getResponseCode());
  const batch=JSON.parse(r.getContentText());
  if(!Array.isArray(batch)) throw Error(table+'応答異常');
  out=out.concat(batch);
  if(batch.length<1000) break;
  offset+=1000;
  if(offset>100000) throw Error('取得件数上限');
 }
 return out;
}
function pqDate_(v,dateOnly) {
 if(!v) return '';
 return Utilities.formatDate(new Date(v),'Asia/Tokyo',dateOnly?'yyyy/MM/dd':'yyyy/MM/dd HH:mm');
}
function pqIndex_(values,col,what){
 const map=new Map();
 for(let i=1;i<values.length;i++){
  const v=String(values[i][col]||'').trim();
  if(!v) continue;
  if(map.has(v)) throw Error(what+' 重複ID: '+v);
  map.set(v,i+1);
 }
 return map;
}
function pqBlank_(values,keys,start){
 const occupied=new Set(keys);
 for(let i=1;i<values.length;i++) if(!occupied.has(i+1) && !values[i].some(v=>v!==''&&v!==null)) return i+1;
 return values.length+1;
}
function pqSet_(updates,sheet,row,col,val){
 updates.push({range:"'"+sheet+"'!"+pqColumn_(col)+row,values:[[val==null?'':val]]});
}
function pqColumn_(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function pqWrite_(updates){
 for(let i=0;i<updates.length;i+=100){
  const chunk=updates.slice(i,i+100);
  const body={valueInputOption:'USER_ENTERED',data:chunk};
  const url='https://sheets.googleapis.com/v4/spreadsheets/'+PQ_SHEET_ID+'/values:batchUpdate';
  const r=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},payload:JSON.stringify(body),muteHttpExceptions:true});
  if(r.getResponseCode()!==200) throw Error('台帳書込失敗 HTTP '+r.getResponseCode()+': '+r.getContentText().slice(0,200));
 }
}
function pqRows_(sheet,lastCol){
 const s=SpreadsheetApp.openById(PQ_SHEET_ID).getSheetByName(sheet);
 if(!s) throw Error('タブなし: '+sheet);
 return {sheet:s,values:s.getRange(1,1,s.getMaxRows(),lastCol).getValues()};
}
function pqSafeRow_(values,row,allowed,what){
 if(row>values.length) throw Error(what+' 空き行不足');
 const v=values[row-1];
 for(let c=0;c<v.length;c++){
  if(allowed.includes(c+1)) continue;
  if(v[c]!=='' && v[c]!==null) throw Error(what+' 行'+row+'に保護対象データあり');
 }
}
function syncPayQuestManual(){
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(1000)) throw Error('同期処理が実行中');
 try{
  const customers=pqFetch_('customers','id,display_name,email,phone,created_at,auth_user_id,customer_rank,risk_status,residency_country_code,residency_status,registered_postal_code,registered_prefecture,registered_city,registered_address_line1,registered_address_line2');
  const quests=pqFetch_('quests','id,quest_code,customer_id,status,created_at,reception_status,reception_reserved_for');
  const items=pqFetch_('quest_items','id,quest_id,item_name,item_url,quantity,procurement_mode,source_key,source_name,requested_unit_price_yen,created_at,source_ward');
  const cu=pqRows_('ユーザー管理',23),qu=pqRows_('QUEST台帳',118),it=pqRows_('QUEST商品明細',27);
  const ci=pqIndex_(cu.values,15,'顧客'),qi=pqIndex_(qu.values,116,'QUEST'),ii=pqIndex_(it.values,23,'商品');
  const userIds=new Set(),userByCustomer=new Map(),questById=new Map(),itemOrdinal=new Map(),updates=[];
  let maxUser=0;
  for(let i=1;i<cu.values.length;i++){
   const uid=String(cu.values[i][0]||'');
   if(uid){if(userIds.has(uid)) throw Error('USER ID重複');userIds.add(uid);const m=/^PQ-U(\d{6})$/.exec(uid);if(m) maxUser=Math.max(maxUser,+m[1])}
  }
  const now=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm');
  const allocated={customer:new Set(ci.values()),quest:new Set(qi.values()),item:new Set(ii.values())};
  function alloc(table,values,idx,id,allowed,what){
   if(idx.has(id)) return idx.get(id);
   let row=pqBlank_(values,allocated[table],2);
   pqSafeRow_(values,row,allowed,what);
   allocated[table].add(row);idx.set(id,row);
   return row;
  }
  for(const c of customers){
   const row=alloc('customer',cu.values,ci,c.id,[1,2,3,4,5,6,7,16,17,18,19,20,21,22,23],'顧客');
   let uid=cu.values[row-1][0];
   if(!uid){do{uid='PQ-U'+String(++maxUser).padStart(6,'0')}while(userIds.has(uid));userIds.add(uid);pqSet_(updates,'ユーザー管理',row,1,uid)}
   userByCustomer.set(c.id,uid);
   const addr=[c.registered_prefecture,c.registered_city,c.registered_address_line1,c.registered_address_line2].filter(Boolean).join('');
   for(const [col,v] of [[2,c.display_name],[3,pqDate_(c.created_at,true)],[4,c.display_name],[5,c.phone],[6,c.registered_postal_code],[7,addr],[16,c.id],[17,c.auth_user_id],[18,c.email],[19,c.customer_rank],[20,c.risk_status],[21,c.residency_country_code],[22,c.residency_status],[23,now]])pqSet_(updates,'ユーザー管理',row,col,v);
  }
  for(const q of quests){
   if(!PQ_STATUS[q.status]) throw Error('未定義QUEST状態 '+q.status);
   if(q.customer_id && !userByCustomer.has(q.customer_id)) throw Error('顧客参照なし '+q.id);
   const row=alloc('quest',qu.values,qi,q.id,[1,2,3,27,117,118,119,120,121],'QUEST');
   questById.set(q.id,q.quest_code);
   for(const [col,v] of [[1,q.quest_code],[2,pqDate_(q.created_at,true)],[3,userByCustomer.get(q.customer_id)||''],[27,PQ_STATUS[q.status]],[117,q.id],[118,q.status],[119,q.reception_status],[120,pqDate_(q.reception_reserved_for,false)],[121,now]])pqSet_(updates,'QUEST台帳',row,col,v);
  }
  for(const item of items){
   const code=questById.get(item.quest_id);if(!code) throw Error('QUEST参照なし '+item.id);
   const ordinal=(itemOrdinal.get(item.quest_id)||0)+1;
   if(ordinal>5) throw Error('QUEST商品5件超過 '+code);
   itemOrdinal.set(item.quest_id,ordinal);
   const row=alloc('item',it.values,ii,item.id,[1,3,4,5,6,7,9,24,25,26,27],'商品');
   const existing=it.values[row-1][2];
   if(existing && Number(existing)!==ordinal) throw Error('商品番号の整合性不一致 '+item.id);
   for(const [col,v] of [[1,code],[3,ordinal],[4,item.procurement_mode==='EC'?'EC':item.procurement_mode==='STORE'?'実店舗':''],[5,item.item_name],[6,item.item_url||item.source_name],[7,item.quantity],[9,item.requested_unit_price_yen],[24,item.id],[25,item.source_key],[26,item.source_ward],[27,now]])pqSet_(updates,'QUEST商品明細',row,col,v);
  }
  // すべての整合性検査に合格した後で書き込み。管理欄・数式は対象外。
  pqWrite_(updates);
  const result='成功｜顧客'+customers.length+'件・QUEST'+quests.length+'件・商品'+items.length+'件';
  pqWrite_([{range:"'同期設定'!B11:B12",values:[[now],[result]]}]);
  return result;
 }catch(e){
  try{pqWrite_([{range:"'同期設定'!B12",values:[['失敗｜'+String(e.message).slice(0,150)]]}])}catch(ignore){}
  throw e;
 }finally{lock.releaseLock()}
}
