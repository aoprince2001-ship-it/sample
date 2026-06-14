// ============================================================
// ひだまり 支援記録 v6.0  Code.gs  ── シンプル・確実動作版
// ============================================================
const SS   = () => SpreadsheetApp.getActiveSpreadsheet();
const SH   = name => { const s=SS().getSheetByName(name); return s||init(name); };
const RECS = 'records', USERS='users', CFG='config', ARC='archive';
const KEEP = {normal:2, important:10, critical:999};

// ─── シート初期化 ───────────────────────────────────────────
function init(name){
  const ss=SS(), sh=ss.insertSheet(name);
  if(name===RECS||name===ARC){
var h=['id','date','user_id','staff',
  'am_work','pm_work',
  'lunch','lunch_amount','lunch_reason','lunch_reason_other',
  'haiben','haiben_time','haiben_amount','haiben_state',
  'shikkin','shikkin_time',
  'event_types',
  'wave_mode','wave_intensity',
  'wave_func','wave_func_detail',
  'wave_behavior','wave_behavior_detail',
  'wave_damage','wave_damage_detail',
  'wave_response','wave_response_detail',
  'wave_result','wave_env','wave_times',
  'safety_type',
'safety_self_type',
  'safety_self_time','safety_self_place','safety_self_content',
  'safety_self_part','safety_self_action','safety_self_after',
  'safety_self_obj_time','safety_self_obj_place','safety_self_obj_name',
  'safety_self_obj_status','safety_self_obj_content','safety_self_obj_after',
  'safety_other_type',
  'safety_other_time','safety_other_place',
  'safety_other_who','safety_other_name',
  'safety_other_part','safety_other_detail',
  'safety_other_action','safety_other_strategy','safety_other_after',
  'safety_obj_time','safety_obj_place','safety_obj_detail',
  'safety_obj_name','safety_obj_status','safety_obj_action',
  'epilepsy_time','epilepsy_duration','epilepsy_scene',
  'epilepsy_detail','epilepsy_action','epilepsy_medicine','epilepsy_after',
  'taichou_time','taichou_place','taichou_symptom',
  'taichou_vital','taichou_action','taichou_medicine','taichou_after',
  'renraku','good_things','sonota',
  'note_ai','case_1lines',
  'importance','archive_due','created_at','updated_at',
 'kosoku_name','kosoku_type','kosoku_time','kosoku_scene',
  'wave_records',
  'taijin_lv','taijin_who','taijin_name','taijin_part','taijin_detail','taijin_action','taijin_medicine','taijin_calm',
  'jishin_lv','jishin_part','jishin_action_type','jishin_detail','jishin_action','jishin_medicine','jishin_calm',
  'taibutsu_owner','taibutsu_name','taibutsu_owner_name','taibutsu_lv','taibutsu_detail','taibutsu_obj'
  ];
    sh.appendRow(h);
    sh.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#2CB67D').setFontColor('#fff');
    sh.setFrozenRows(1);
  } else if(name===USERS){
    sh.appendRow(['user_id','name','active']);
    sh.appendRow(['U001','青木 凜太','TRUE']);
    sh.appendRow(['U002','村上 さくら','TRUE']);
    sh.appendRow(['U003','橋本 柊哉','TRUE']);
  } else if(name===CFG){
    sh.appendRow(['key','value']);
  　 [['admin_pw','1234'],
     ['tmpl_base','午前は{am_work}に取り組まれました。午後は{pm_work}をして過ごされています。'],
     ['tmpl_lunch_full','昼食は完食されています。'],
     ['tmpl_lunch_amount_low','ほとんど食べられました。'],
     ['tmpl_lunch_amount_mid','半分ほど召し上がられました。'],
     ['tmpl_lunch_amount_less','3割ほど召し上がられました。'],
     ['tmpl_lunch_amount_none','ほとんど食べられませんでした。'],
     ['tmpl_lunch_hungry','帰宅後、空腹を感じる場合もあるかと思われますので、ご様子を見ていただければ幸いです。'],
     ['tmpl_lunch_reason_full','お腹がいっぱいであった可能性が考えられます。'],
     ['tmpl_lunch_reason_taste','嗜好やその時の気分による影響もあった様子です。'],
     ['tmpl_lunch_reason_health','体調面の影響があった可能性が考えられます。'],
     ['tmpl_lunch_reason_self','{reason}であると本人さんが自身で伝えてくださいました。'],
     ['tmpl_lunch_reason_other','{reason}と考えられます。'],
     ['tmpl_wave','本日、{time}頃、{behavior}の様子が見られ{intensity}ある場面がありました。{func}のため、{response}にて対応しております。その後は{result}。'],
     ['tmpl_intensity_low','少し気持ちの波が'],
     ['tmpl_intensity_mid','気持ちの高まりが'],
     ['tmpl_intensity_high','強い気持ちの高まりが'],
     ['tmpl_result_calm','気持ちの切り替えができ落ち着いて過ごされています。'],
     ['tmpl_result_cont','まだ少しモヤモヤした様子で過ごされました。'],
     ['tmpl_safety_injury','本日、{time}頃、{place}にて{content}。{part}に{medicine}{action}を行い、その後は大きな変化なく過ごされています。'],
     ['tmpl_safety_obj','本日、{time}頃、{place}にて{obj_name}に{status}が見られました。{after}を行っております。'],
     ['tmpl_safety_other','本日、周囲との関わりの中で{part}に影響が見られました。{action}を行い、その後は{after}様子で過ごされています。\nご利用中の出来事によりご心配をおかけし申し訳ございません。'],
     ['tmpl_epilepsy','本日、{time}頃にてんかん発作と思われる様子が見られました。発作時は{detail}、{action}を行い、{medicine}その後は{after}様子で過ごされています。'],
     ['tmpl_taichou','本日、{time}頃、{symptom}が見られました。バイタルは{vital}。{action}{medicine}を行い、その後の体調は{after}。'],
     ['tmpl_kosoku_on','本日、{time}頃、安全面への配慮から{scene}のため{name}を使用しています。'],
     ['tmpl_kosoku_off','本日、{time}頃に{scene}のため{name}を一時的に外させていただいております。'],
    ['tmpl_no_event','本日は大きな変化なく、終日穏やかに過ごされています😊'],
     ['tmpl_haiben','本日、{time}頃、排便がありました。（便の状態：{state}　便の量：{amount}）'],
     ['tmpl_shikkin_reason_shikkin','排泄のタイミングが合わず'],
     ['tmpl_shikkin_reason_more','排泄に伴う衣類の汚れが見られ'],
     ['tmpl_shikkin_reason_dirt','衣類の汚れが見られ'],
     ['tmpl_shikkin_reason_temp','気候による影響があり'],
     ['tmpl_shikkin_reason_self','ご本人様の希望により'],
     ['tmpl_shikkin','{reason}{clothes}のお着替えをしています。'],
     ['staff','田中,川島,木村,中川,森田,石川,岡田,渡辺'],
     ['am','ボルト作業,ビス作業,ペグ作業,手伝い,その他'],
     ['am_note','ボルト作業:午前はボルト作業に取り組まれました。|ビス作業:午前はビス作業に取り組まれました。|ペグ作業:午前はペグ作業に取り組まれました。|手伝い:午前はお手伝いをして過ごされています。|その他:午前は{detail}をして過ごされています。'],
     ['pm','レクリエーション参加,自由時間,お手伝い,横になる,その他'],
     ['pm_note','レクリエーション参加:午後はレクリエーションに参加して過ごされています。|自由時間:午後は自由時間でのんびりと過ごされています。|お手伝い:午後はお手伝いをして過ごされています。|横になる:午後は横になって身体を休める時間を取りながら過ごされています。|その他:午後は{detail}をして過ごされています。'],
     ['lunch_reason','おなかいっぱい,体調不良,好みでなかった,本人の希望,その他'],
     ['day_status','元気に過ごされた,気持ちの波あり,体調不良あり,けが・傷あり,てんかん発作あり,その他'],
     ['wave_expr','涙,大声,硬直,自傷,対物,対人'],
     ['wave_state','イライラ,眠たい,不快感,悲しそう,寂しそう,希望通りにならなかった,職員に見てほしかった'],
     ['wave_after','自身で切り替えられた,横になり休まれた,もやもやしている様子,別室で切り替えられた,その他'],
     ['timeslots','9時頃,10時頃,11時頃,12時頃,13時頃,14時頃,15時頃,16時頃,来所時,帰宅前,終日'],
     ['renraku','連絡物あり,きらりほっと報告,質問,その他'],
     ['incident_types','転倒,大声,離席,自傷,他者への攻撃,物品破損,飛び出し,過食,拒食,薬の拒否,排泄問題,その他'],
     ['locations','作業室,食堂,トイレ,廊下,玄関,休憩室,屋外,送迎車内,その他'],
     ['case_results','落ち着かれた,継続対応中,保護者に連絡,看護師に連絡,管理者に報告,記録のみ,その他'],
     ['wave_env_options','周囲が騒がしかった,見慣れない職員がいた,初めての活動があった,好きではない活動があった,担当職員の休みがあった,仲の良い利用者さんが休みだった,周囲の利用者さんからの影響があった,園は普段と変わりない環境だった,その他'],
['wave_func_options','注目,要求,逃避,自己刺激,分析不能,その他'],
['note_b_map','大声:大きな声が見られました|泣き:涙が見られました|硬直:体が固まる様子が見られました|唾吐き:口元での行動が見られました|暴言:強い言葉が見られました|噛む:口での強い接触が見られました|暴力:周囲に向かう強い関わりが見られました|落ち着きがない:落ち着かない様子が見られました|物を投げる:物へ向かい手に取る様子が見られました|物を壊す:物の扱いに強さが見られました|離席:その場を離れる様子が見られました|脱衣:服を気にする様子がありました|パニック:混乱している様子が見られました|威嚇:気持ちの高まりを周囲に向ける様子がありました|座り込み・寝転び:その場に留まり動き出しにくい様子がありました|拒否:声掛けに対して受け入れにくい様子が見られました|口にいれる:物を口へ持って行く様子が見られました|過度な接触:他者との距離が近くなる様子が見られました|水遊び:水に関心が向く様子が見られました|物盗み:物に関心が向き持ち帰ろうとする様子が見られました|強い緊張:強い緊張が見られました'],
['note_a_map','注目:人との関わりを求める様子があり|要求:思い通りにならない場面があり|逃避:難しさを感じる場面があり|感覚:感覚面での調整が難しい様子があり|他者影響:周囲の様子に影響を受ける場面があり|内的要因:体調や気分の影響もあり|不明:状況を踏まえ観察を行いましたが明確な要因の特定には至りませんでした|見通しへの不安:見通しが持ちにくい場面があり'],
['note_c_map','気持ちが落ち着いたタイミングで声かけ:声かけ|周囲との距離を確保:距離の確保|別室移動:別室での対応|服薬:服薬|他のことで気持ちをそらす:他の活動への切り替え|良い行動・正しい行動の時に褒めた:適切な行動への声かけ|代替行動を教育した:代替行動の提示|環境調整:環境調整|見守り:見守り'],
    ].forEach(function(r){ sh.appendRow(r); });
  }
  return sh;
}

// ─── doGet ──────────────────────────────────────────────────
function doGet(){
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ひだまり支援記録')
    .addMetaTag('viewport','width=device-width,initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// ─── Config ─────────────────────────────────────────────────
function getConfig(){
  const d=SH(CFG).getDataRange().getValues(), c={};
  for(let i=1;i<d.length;i++) if(d[i][0]) c[d[i][0]]=String(d[i][1]);
  return c;
}
function setConfig(k,v){
  const sh=SH(CFG), d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++) if(d[i][0]===k){sh.getRange(i+1,2).setValue(v);return;}
  sh.appendRow([k,v]);
}
function saveConfigs(obj){ Object.keys(obj).forEach(k=>setConfig(k,obj[k])); return {ok:true}; }

// ─── Users ──────────────────────────────────────────────────
function getUsers(){
  const d=SH(USERS).getDataRange().getValues();
  return d.slice(1).filter(r=>String(r[2])!=='FALSE')
    .map(r=>({id:String(r[0]),name:String(r[1])}));
}
function addUser(name, yomi){
  const id='U'+Date.now().toString().slice(-6);
  SH(USERS).appendRow([id,name,'',yomi||'','TRUE']);
  CacheService.getScriptCache().remove('initialData');
  return {ok:true,id,name,yomi:yomi||''};
}

// ─── 日付変換 ────────────────────────────────────────────────
function toDate(v){
  if(!v) return '';
  if(v instanceof Date){
    return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
  }
  return String(v).slice(0,10);
}

// ─── 重要度 ──────────────────────────────────────────────────
function importance(rec){
  var t=JSON.stringify(rec);
  if(['救急','入院','骨折','権利侵害','虐待','通報','警察'].some(function(k){ return t.includes(k); })) return 'critical';
  if(['てんかん','発作','嘔吐','発熱','受診'].some(function(k){ return t.includes(k); })||parseInt(rec.wave_intensity||0)>=4) return 'important';
  return 'normal';
}

// ─── 記録保存 ────────────────────────────────────────────────
function saveRecord(rec){
  const sh=SH(RECS), now=new Date().toISOString();
  const id='R'+Date.now().toString(36).toUpperCase();
  if(!Array.isArray(rec.cases)) rec.cases=[];
  const imp=importance(rec);
  const due=new Date(rec.date||now);
  due.setFullYear(due.getFullYear()+KEEP[imp]);

sh.appendRow([
  id, rec.date||'', rec.user_id||'', rec.staff||'',
  rec.am_work||'', rec.pm_work||'',
  rec.lunch||'', rec.lunch_amount||'', rec.lunch_reason||'', rec.lunch_reason_other||'',
  rec.haiben||'', rec.haiben_time||'', rec.haiben_amount||'', rec.haiben_state||'',
  rec.shikkin||'', rec.shikkin_time||'', rec.shikkin_clothes||'', rec.shikkin_reason||'',
  rec.event_types||'',
  rec.wave_mode||'', rec.wave_intensity||'',
  rec.wave_func||'', rec.wave_func_detail||'',
  rec.wave_behavior||'', rec.wave_behavior_detail||'',
  rec.wave_damage||'', rec.wave_damage_detail||'',
  rec.wave_response||'', rec.wave_response_detail||'',
  rec.wave_result||'', rec.wave_env||'', rec.wave_times||'',
  rec.safety_type||'',
  rec.safety_self_type||'',
  rec.safety_self_time||'', rec.safety_self_place||'', rec.safety_self_content||'',
  rec.safety_self_part||'', rec.safety_self_action||'', rec.safety_self_after||'', rec.safety_self_medicine||'',
  rec.safety_self_obj_time||'', rec.safety_self_obj_place||'', rec.safety_self_obj_name||'',
  rec.safety_self_obj_status||'', rec.safety_self_obj_content||'', rec.safety_self_obj_after||'',
  rec.safety_other_type||'',
  rec.safety_other_time||'', rec.safety_other_place||'',
  rec.safety_other_who||'', rec.safety_other_name||'',
  rec.safety_other_part||'', rec.safety_other_detail||'',
  rec.safety_other_action||'', rec.safety_other_strategy||'', rec.safety_other_after||'',
  rec.safety_obj_time||'', rec.safety_obj_place||'', rec.safety_obj_detail||'',
  rec.safety_obj_name||'', rec.safety_obj_status||'', rec.safety_obj_action||'',
  rec.epilepsy_time||'', rec.epilepsy_duration||'', rec.epilepsy_scene||'',
  rec.epilepsy_detail||'', rec.epilepsy_action||'', rec.epilepsy_medicine||'', rec.epilepsy_after||'',
  rec.taichou_time||'', rec.taichou_place||'', rec.taichou_symptom||'',
  rec.taichou_vital||'', rec.taichou_action||'', rec.taichou_medicine||'', rec.taichou_after||'',
  rec.renraku||'', rec.good_things||'', rec.sonota||'', rec.katei_renraku||'',
  '', '',
  imp, due.toISOString().split('T')[0], now, now,
  rec.kosoku_name||'', rec.kosoku_type||'', rec.kosoku_time||'', rec.kosoku_scene||'',
  rec.wave_records||'',
  rec.taijin_lv||'', rec.taijin_who||'', rec.taijin_name||'', rec.taijin_part||'', rec.taijin_detail||'', rec.taijin_action||'', rec.taijin_medicine||'', rec.taijin_calm||'',
  rec.jishin_lv||'', rec.jishin_part||'', rec.jishin_detail||'', rec.jishin_action_type||'', rec.jishin_action||'', rec.jishin_medicine||'', rec.jishin_calm||'',
  rec.taibutsu_owner||'', rec.taibutsu_name||'', rec.taibutsu_owner_name||'', rec.taibutsu_lv||'', rec.taibutsu_detail||'', rec.taibutsu_obj||'',
  rec.gokuin_type||'',
  rec.ensen_symptom||'', rec.ensen_situation||'', rec.ensen_cause||'', rec.ensen_cause_detail||'',
  rec.ensen_action||'', rec.ensen_action_detail||'', rec.ensen_after||'', rec.ensen_after_memo||'',
  rec.ensen_plan||'', rec.ensen_plan_detail||'',
  rec.gokuin_object||'', rec.gokuin_place||'', rec.gokuin_situation||'',
  rec.gokuin_cause||'', rec.gokuin_cause_detail||'',
  rec.gokuin_action||'', rec.gokuin_action_detail||'',
  rec.gokuin_after||'', rec.gokuin_after_memo||'',
  rec.gokuin_plan||'', rec.gokuin_plan_detail||'',
  rec.ensen_lv||'', rec.ensen_time||'', rec.ensen_place||'', rec.ensen_scene||'', rec.ensen_food||'',
  rec.gokuin_danger||'', rec.gokuin_swallow||'', rec.gokuin_lv||'',
  rec.gokuin_time||'', rec.gokuin_scene||'',
  rec.shikkin_before||'', rec.shikkin_after_detail||'',
  rec.safety_self_medicine||'', rec.safety_other_medicine||''
]);
  return {ok:true, id, importance:imp};
}

// ─── フィールド更新 ──────────────────────────────────────────
function updateField(id, field, value){
  const sh=SH(RECS), d=sh.getDataRange().getValues(), h=d[0];
  const col=h.indexOf(field); if(col<0) return {ok:false};
  for(let i=1;i<d.length;i++){
    if(String(d[i][0])!==String(id)) continue;
    sh.getRange(i+1,col+1).setValue(value);
    sh.getRange(i+1,h.indexOf('updated_at')+1).setValue(new Date().toISOString());
    return {ok:true};
  }
  return {ok:false};
}

// ─── 記録更新 ────────────────────────────────────────────────
function updateRecord(id, updates){
  const sh=SH(RECS), d=sh.getDataRange().getValues(), h=d[0];
  for(let i=1;i<d.length;i++){
    if(String(d[i][0])!==String(id)) continue;
    Object.keys(updates).forEach(k=>{const c=h.indexOf(k);if(c>=0)sh.getRange(i+1,c+1).setValue(updates[k]);});
    sh.getRange(i+1,h.indexOf('updated_at')+1).setValue(new Date().toISOString());
    return {ok:true};
  }
  return {ok:false};
}

// ─── 行パース ────────────────────────────────────────────────
function parseRow(h, row){
  const o={};
  h.forEach((k,j)=>{ o[k]=(row[j] instanceof Date)?toDate(row[j]):row[j]; });
  try{ o.cases=JSON.parse(o.cases_json||'[]'); }catch(e){ o.cases=[]; }
  try{ o.case_1lines_arr=JSON.parse(o.case_1lines||'[]'); }catch(e){ o.case_1lines_arr=[]; }
  o.hasWave=!!(o.wave_time||String(o.day_status||'').includes('気持ちの波'));
  return o;
}

// ─── 記録取得 ────────────────────────────────────────────────
function getRecords(opts){
  opts=opts||{};
  const sh=SH(RECS);
  if(sh.getLastRow()<=1) return [];
  const d=sh.getDataRange().getValues(), h=d[0];
  let rows=d.slice(1).map(r=>parseRow(h,r));
  if(opts.date)      rows=rows.filter(r=>toDate(r.date)===opts.date);
  if(opts.user_id)   rows=rows.filter(r=>String(r.user_id)===String(opts.user_id));
  if(opts.date_from) rows=rows.filter(r=>toDate(r.date)>=opts.date_from);
  if(opts.date_to)   rows=rows.filter(r=>toDate(r.date)<=opts.date_to);
  if(opts.keyword){const kw=opts.keyword.toLowerCase();rows=rows.filter(r=>JSON.stringify(r).toLowerCase().includes(kw));}
  rows.sort((a,b)=>toDate(b.date).localeCompare(toDate(a.date)));
  return rows.slice(0,opts.limit||200);
}

function getRecordById(id){
  return getRecords({limit:9999}).find(r=>String(r.id)===String(id))||null;
}

// ─── 印刷データ取得 ──────────────────────────────────────────
function getPrintData(opts){
  return getRecords(opts).map(function(r){
    return {
      id:String(r.id||''), date:String(r.date||''), user_id:String(r.user_id||''),
      staff:String(r.staff||''), am_work:String(r.am_work||''), pm_work:String(r.pm_work||''),
      lunch:String(r.lunch||''), lunch_amount:String(r.lunch_amount||''),
      lunch_reason:String(r.lunch_reason||''), lunch_reason_other:String(r.lunch_reason_other||''),
      haiben:String(r.haiben||''), haiben_state:String(r.haiben_state||''), haiben_amount:String(r.haiben_amount||''),
      shikkin:String(r.shikkin||''), shikkin_reason:String(r.shikkin_reason||''), shikkin_clothes:String(r.shikkin_clothes||''),
      shikkin_before:String(r.shikkin_before||''), shikkin_after_detail:String(r.shikkin_after_detail||''),
      event_types:String(r.event_types||''),
      wave_mode:String(r.wave_mode||''), wave_records:String(r.wave_records||''),
      wave_func:String(r.wave_func||''), wave_func_detail:String(r.wave_func_detail||''),
      wave_behavior:String(r.wave_behavior||''), wave_behavior_detail:String(r.wave_behavior_detail||''),
      wave_response:String(r.wave_response||''), wave_response_detail:String(r.wave_response_detail||''),
      wave_result:String(r.wave_result||''), wave_times:String(r.wave_times||''),
      wave_intensity:parseInt(r.wave_intensity||0),
      safety_type:String(r.safety_type||''),
      safety_self_type:String(r.safety_self_type||''),
      safety_self_time:String(r.safety_self_time||''), safety_self_place:String(r.safety_self_place||''),
      safety_self_content:String(r.safety_self_content||''), safety_self_part:String(r.safety_self_part||''),
      safety_self_action:String(r.safety_self_action||''), safety_self_after:String(r.safety_self_after||''),
      safety_self_medicine:String(r.safety_self_medicine||''),
      safety_self_obj_time:String(r.safety_self_obj_time||''), safety_self_obj_place:String(r.safety_self_obj_place||''),
      safety_self_obj_name:String(r.safety_self_obj_name||''), safety_self_obj_status:String(r.safety_self_obj_status||''),
      safety_self_obj_content:String(r.safety_self_obj_content||''), safety_self_obj_after:String(r.safety_self_obj_after||''),
      safety_other_type:String(r.safety_other_type||''),
      safety_other_time:String(r.safety_other_time||''), safety_other_place:String(r.safety_other_place||''),
      safety_other_who:String(r.safety_other_who||''), safety_other_name:String(r.safety_other_name||''),
      safety_other_part:String(r.safety_other_part||''), safety_other_detail:String(r.safety_other_detail||''),
      safety_other_action:String(r.safety_other_action||''), safety_other_strategy:String(r.safety_other_strategy||''),
      safety_other_after:String(r.safety_other_after||''), safety_other_medicine:String(r.safety_other_medicine||''),
      safety_obj_time:String(r.safety_obj_time||''), safety_obj_place:String(r.safety_obj_place||''),
      safety_obj_detail:String(r.safety_obj_detail||''), safety_obj_name:String(r.safety_obj_name||''),
      safety_obj_status:String(r.safety_obj_status||''), safety_obj_action:String(r.safety_obj_action||''),
      epilepsy_time:String(r.epilepsy_time||''), epilepsy_duration:String(r.epilepsy_duration||''),
      epilepsy_scene:String(r.epilepsy_scene||''), epilepsy_detail:String(r.epilepsy_detail||''),
      epilepsy_action:String(r.epilepsy_action||''), epilepsy_medicine:String(r.epilepsy_medicine||''),
      epilepsy_after:String(r.epilepsy_after||''),
      taichou_time:String(r.taichou_time||''), taichou_place:String(r.taichou_place||''),
      taichou_symptom:String(r.taichou_symptom||''), taichou_vital:String(r.taichou_vital||''),
      taichou_action:String(r.taichou_action||''), taichou_medicine:String(r.taichou_medicine||''),
      taichou_after:String(r.taichou_after||''),
      kosoku_name:String(r.kosoku_name||''), kosoku_type:String(r.kosoku_type||''),
      kosoku_time:String(r.kosoku_time||''), kosoku_scene:String(r.kosoku_scene||''),
      renraku:String(r.renraku||''), good_things:String(r.good_things||''),
      sonota:String(r.sonota||''), katei_renraku:String(r.katei_renraku||''),
      note_ai:String(r.note_ai||''),
      gokuin_type:String(r.gokuin_type||''),
      ensen_lv:String(r.ensen_lv||''), ensen_time:String(r.ensen_time||''),
      ensen_place:String(r.ensen_place||''), ensen_scene:String(r.ensen_scene||''),
      ensen_food:String(r.ensen_food||''), ensen_symptom:String(r.ensen_symptom||''),
      ensen_cause:String(r.ensen_cause||''), ensen_action:String(r.ensen_action||''),
      ensen_after:String(r.ensen_after||''), ensen_plan:String(r.ensen_plan||''),
      ensen_situation:String(r.ensen_situation||''), ensen_cause_detail:String(r.ensen_cause_detail||''),
      ensen_action_detail:String(r.ensen_action_detail||''), ensen_after_memo:String(r.ensen_after_memo||''),
      ensen_plan_detail:String(r.ensen_plan_detail||''),
      gokuin_lv:String(r.gokuin_lv||''), gokuin_time:String(r.gokuin_time||''),
      gokuin_object:String(r.gokuin_object||''), gokuin_place:String(r.gokuin_place||''),
      gokuin_scene:String(r.gokuin_scene||''), gokuin_situation:String(r.gokuin_situation||''),
      gokuin_cause:String(r.gokuin_cause||''), gokuin_action:String(r.gokuin_action||''),
      gokuin_after:String(r.gokuin_after||''), gokuin_plan:String(r.gokuin_plan||''),
      gokuin_danger:String(r.gokuin_danger||''), gokuin_swallow:String(r.gokuin_swallow||''),
      gokuin_cause_detail:String(r.gokuin_cause_detail||''), gokuin_action_detail:String(r.gokuin_action_detail||''),
      gokuin_after_memo:String(r.gokuin_after_memo||''), gokuin_plan_detail:String(r.gokuin_plan_detail||'')
    };
  });
}

// ─── ABAデータ取得 ───────────────────────────────────────────
function getAbaData(opts){
  var rows=getRecords({user_id:opts.user_id,date_from:opts.date_from,date_to:opts.date_to,limit:500});
  var result=[];
  rows.forEach(function(r){
    var hasWave = r.event_types && r.event_types.indexOf('気持ちの波あり') !== -1;
    var hasWazato = r.shikkin === 'あり' && r.shikkin_reason && r.shikkin_reason.indexOf('わざと・意図的な可能性あり') !== -1;
    if(!hasWave && !hasWazato) return;
    if(hasWave){
      if(r.wave_mode==='単発'||r.wave_mode===''){
        var wRecs=[];
        try{
          var parsed=JSON.parse(String(r.wave_records||'[]'));
          wRecs=Array.isArray(parsed)?parsed:[];
        }catch(e){ wRecs=[]; }        wRecs.forEach(function(rec){
          result.push({
            record_id:r.id, date:r.date, staff:r.staff||'',
            antecedent:rec.func||'', antecedent_detail:rec.func_detail||'',
            behavior:rec.behavior||'', behavior_detail:rec.behavior_detail||'',
            consequence:rec.response||'', consequence_detail:rec.response_detail||'',
            result:rec.result||'', cont_detail:rec.cont_detail||'',
            intensity:parseInt(rec.intensity||0),
            time:rec.time||'', place:rec.place||'', scene:rec.scene||'',
            damage:rec.damage||'',
            taijin_lv:rec.taijin_lv||'', taijin_who:rec.taijin_who||'',
            taijin_part:rec.taijin_part||'', taijin_detail:rec.taijin_detail||'',
            jishin_lv:rec.jishin_lv||'', jishin_action_type:rec.jishin_action_type||'',
            jishin_part:rec.jishin_part||'', jishin_detail:rec.jishin_detail||'',
            taibutsu_owner:rec.taibutsu_owner||'', taibutsu_name:rec.taibutsu_name||'',
            taibutsu_lv:rec.taibutsu_lv||''
          });
        });
      } else {
        result.push({
          record_id:r.id, date:r.date, staff:r.staff||'',
          antecedent:r.wave_func||'', antecedent_detail:r.wave_func_detail||'',
          behavior:r.wave_behavior||'', behavior_detail:r.wave_behavior_detail||'',
          consequence:r.wave_response||'', consequence_detail:r.wave_response_detail||'',
          result:r.wave_result||'', cont_detail:'',
          intensity:parseInt(r.wave_intensity||0),
          time:r.wave_times||'', place:'', scene:'',
          damage:''
        });
      }
    }
    if(hasWazato && !hasWave){
      result.push({
        record_id:r.id, date:r.date, staff:r.staff||'',
        antecedent:'わざと・意図的な可能性あり', antecedent_detail:r.shikkin_before||'',
        behavior:'失禁・漏れ', behavior_detail:r.shikkin_reason||'',
        consequence:'更衣対応', consequence_detail:r.shikkin_after_detail||'',
        result:'', cont_detail:'',
        intensity:1,
        time:r.shikkin_time||'', place:'', scene:'',
        damage:''
      });
    }
  });
  result.sort(function(a,b){ return b.date.localeCompare(a.date); });
  return result;
}
// ─── 認証・初期データ ────────────────────────────────────────
function checkPw(pw){ return pw===(getConfig().admin_pw||'1234'); }

function getInitialData(){
  Logger.log('start');
  CacheService.getScriptCache().removeAll(['initialData']);
  // キャッシュから取得を試みる
  const cache = CacheService.getScriptCache();
  const cached = cache.get('initialData');
  if(cached){
    try{ return JSON.parse(cached); }catch(e){}
  }

  // キャッシュがなければスプレッドシートから取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let cfg = {};
  try {
    const cs = ss.getSheetByName('config');
    if(cs){
      const d = cs.getDataRange().getValues();
      for(let i=1;i<d.length;i++) if(d[i][0]) cfg[d[i][0]]=String(d[i][1]);
    } else {
      init('config');
      cfg = getConfig();
    }
  } catch(e){ cfg = {}; }

  let users = [];
  try {
    const us = ss.getSheetByName('users');
    if(us){
      const d = us.getDataRange().getValues();
      users = d.slice(1)
        .filter(r=>String(r[4])!=='FALSE')
        .map(r=>({id:String(r[0]),name:String(r[1]),yomi:String(r[3]||r[2]||'')}));    } else {
      init('users');
      users = getUsers();
    }
  } catch(e){ users = []; }

  if(!ss.getSheetByName('records')) init('records');

  const p = k => (cfg[k]||'').split(',').map(s=>s.trim()).filter(Boolean);
Logger.log(JSON.stringify(cfg));
var parseMap = function(k) {
    var val = cfg[k]||'';
    var map = {};
    val.split('|').forEach(function(pair) {
      var idx = pair.indexOf(':');
      if(idx > 0) map[pair.slice(0,idx).trim()] = pair.slice(idx+1).trim();
    });
    return map;
  };
  var result = {
    users: users,
    staff:          p('staff'),
    am:             p('am'),
    pm:             p('pm'),
    lunch_reason:   p('lunch_reason'),
    day_status:     p('day_status'),
    wave_expr:      p('wave_expr'),
    wave_state:     p('wave_state'),
    wave_after:     p('wave_after'),
    timeslots:      p('timeslots'),
    renraku:        p('renraku'),
    incident_types: p('incident_types'),
    locations:      p('locations'),
    case_results:   p('case_results'),
    wave_env_options: p('wave_env_options'),
    wave_func_options: p('wave_func_options'),
note_b_map: parseMap('note_b_map'),
    note_a_map: parseMap('note_a_map'),
    note_c_map: parseMap('note_c_map'),
    am_note_map: parseMap('am_note'),
    pm_note_map: parseMap('pm_note'),
    tmpl: {
      base:               cfg['tmpl_base']||'',
      lunch_full:         cfg['tmpl_lunch_full']||'',
      lunch_amount_low:   cfg['tmpl_lunch_amount_low']||'',
      lunch_amount_mid:   cfg['tmpl_lunch_amount_mid']||'',
      lunch_amount_less:  cfg['tmpl_lunch_amount_less']||'',
      lunch_amount_none:  cfg['tmpl_lunch_amount_none']||'',
      lunch_hungry:       cfg['tmpl_lunch_hungry']||'',
      lunch_reason_full:  cfg['tmpl_lunch_reason_full']||'',
      lunch_reason_taste: cfg['tmpl_lunch_reason_taste']||'',
      lunch_reason_health:cfg['tmpl_lunch_reason_health']||'',
      lunch_reason_self:  cfg['tmpl_lunch_reason_self']||'',
      lunch_reason_other: cfg['tmpl_lunch_reason_other']||'',
      wave:               cfg['tmpl_wave']||'',
      intensity_low:      cfg['tmpl_intensity_low']||'',
      intensity_mid:      cfg['tmpl_intensity_mid']||'',
      intensity_high:     cfg['tmpl_intensity_high']||'',
      result_calm:        cfg['tmpl_result_calm']||'',
      result_cont:        cfg['tmpl_result_cont']||'',
      safety_injury:      cfg['tmpl_safety_injury']||'',
      safety_obj:         cfg['tmpl_safety_obj']||'',
      safety_other:       cfg['tmpl_safety_other']||'',
      epilepsy:           cfg['tmpl_epilepsy']||'',
      taichou:            cfg['tmpl_taichou']||'',
 　　　kosoku_on:          cfg['tmpl_kosoku_on']||'',
      kosoku_off:         cfg['tmpl_kosoku_off']||'',
      no_event:           cfg['tmpl_no_event']||'',
      haiben:             cfg['tmpl_haiben']||'',
      shikkin:            cfg['tmpl_shikkin']||'',
      shikkin_reason_shikkin: cfg['tmpl_shikkin_reason_shikkin']||'',
      shikkin_reason_more:    cfg['tmpl_shikkin_reason_more']||'',
      shikkin_reason_dirt:    cfg['tmpl_shikkin_reason_dirt']||'',
      shikkin_reason_temp:    cfg['tmpl_shikkin_reason_temp']||'',
      shikkin_reason_self:    cfg['tmpl_shikkin_reason_self']||'',
    },
  };

  // 6時間キャッシュに保存
  try{
    cache.put('initialData', JSON.stringify(result), 21600);
  }catch(e){}

  return result;
}
function clearCache(){
  CacheService.getScriptCache().remove('initialData');
}

// ─── Gemini AI生成 ──────────────────────────────────────────
function callGemini(prompt) {
var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=' + getGeminiKey();
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {'Content-Type': 'application/json'},
    payload: JSON.stringify({
      contents: [{parts: [{text: prompt}]}],
      generationConfig: {maxOutputTokens: 1000}
    }),
    muteHttpExceptions: true
  });
  var d = JSON.parse(res.getContentText());
  var text = '';
  try { text = d.candidates[0].content.parts[0].text; } catch(e) { text = ''; }
  return {text: text};
}
function getGeminiKey() {
  return getConfig()['gemini_api_key'] || '';
}

function testSave(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('records');
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange('A1').setValue('シート名:'+sh.getName()+' 行数:'+sh.getLastRow());
}
// ─── 設定更新 ──────────────────────────────────────────────
function updateConfig(key, value){
  var sh = SH(CFG);
  var d = sh.getDataRange().getValues();
  for(var i=1; i<d.length; i++){
    if(d[i][0]===key){
      sh.getRange(i+1,2).setValue(value);
      CacheService.getScriptCache().remove('initialData');
      return {ok:true};
    }
  }
  sh.appendRow([key,value]);
  CacheService.getScriptCache().remove('initialData');
  return {ok:true};
}

function deleteUser(id){
  var sh = SH(USERS);
  var d = sh.getDataRange().getValues();
  for(var i=1; i<d.length; i++){
    if(d[i][0]===id){
      sh.deleteRow(i+1);
      CacheService.getScriptCache().remove('initialData');
      return {ok:true};
    }
  }
  return {ok:false};
}

function checkNoteMap() {
  var cfg = getConfig();
  Logger.log(cfg['note_b_map']);
}

function checkMap() {
  Logger.log(getConfig()['note_b_map']);
}