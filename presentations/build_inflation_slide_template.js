const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
pres.title = 'インフレスライドの計算方法と簡略化について';
const F = 'Meiryo';
const C = { navy:'1F3A4D', navy2:'2E5470', amber:'E08E2B', light:'EEF2F5', ph:'8A97A3', text:'2B2B2B', white:'FFFFFF', line:'C9D3DB' };
const W = 13.333;

function title(s, t, sub) {
  s.background = { color: C.white };
  s.addText(t, { x:0.6, y:0.4, w:12.1, h:0.8, fontFace:F, fontSize:30, bold:true, color:C.navy, margin:0, isTextBox:true });
  if (sub) s.addText(sub, { x:0.6, y:1.15, w:12.1, h:0.4, fontFace:F, fontSize:14, color:C.ph, margin:0, isTextBox:true });
}
function footer(s, n) {
  s.addText(String(n), { x:12.2, y:7.0, w:0.6, h:0.3, fontFace:F, fontSize:10, color:C.ph, align:'right', margin:0, isTextBox:true });
}
const ph = (t, o={}) => ({ text:t, options:{ color:C.ph, ...o } });
function circleNum(s, n, x, y, d=0.55, fill=C.amber) {
  s.addShape(pres.shapes.OVAL, { x, y, w:d, h:d, fill:{color:fill}, line:{color:fill} });
  s.addText(String(n), { x, y, w:d, h:d, fontFace:F, fontSize:16, bold:true, color:C.white, align:'center', valign:'middle', margin:0, isTextBox:true });
}
function card(s, x, y, w, h, fill=C.light) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill:{color:fill}, line:{color:fill}, rectRadius:0.08 });
}
function section(no, t, desc) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addText(String(no).padStart(2,'0'), { x:0.9, y:2.2, w:3, h:1.4, fontFace:'Arial', fontSize:80, bold:true, color:C.amber, margin:0, isTextBox:true });
  s.addText(t, { x:0.9, y:3.7, w:11.5, h:0.9, fontFace:F, fontSize:34, bold:true, color:C.white, margin:0, isTextBox:true });
  s.addText(desc, { x:0.9, y:4.6, w:11.5, h:0.5, fontFace:F, fontSize:16, color:'B9C8D4', margin:0, isTextBox:true });
  return s;
}
let n = 1;

// 1 表紙
{ const s = pres.addSlide(); s.background = { color:C.navy };
  // motif: rising steps (price increase)
  [1.2,1.9,2.7,3.6].forEach((h,i)=> s.addShape(pres.shapes.RECTANGLE,{ x:9.4+i*0.85, y:6.6-h, w:0.65, h, fill:{color:i===3?C.amber:C.navy2}, line:{color:i===3?C.amber:C.navy2} }));
  s.addText('インフレスライドの計算方法と\n簡略化について', { x:0.9, y:1.7, w:8.5, h:2.0, fontFace:F, fontSize:36, bold:true, color:C.white, margin:0, isTextBox:true, valign:'top' });
  s.addText('【会議名】打合せ資料', { x:0.9, y:3.9, w:8, h:0.5, fontFace:F, fontSize:18, color:'B9C8D4', margin:0, isTextBox:true });
  s.addText('【20XX年XX月XX日】　【部署名・作成者名】', { x:0.9, y:5.8, w:8, h:0.4, fontFace:F, fontSize:14, color:'B9C8D4', margin:0, isTextBox:true });
  s.addNotes('表紙。会議名・日付・作成者を記入。');
}
n++;

// 2 アジェンダ
{ const s = pres.addSlide(); title(s,'本日のアジェンダ');
  const items = [
    ['会議の目的・背景','本日決めたいこと／共有したいこと'],
    ['インフレスライドとは','制度の概要・適用要件・手続きの流れ'],
    ['計算方法','計算式・手順・計算例'],
    ['計算方法の簡略化','現行の課題・簡略化案の比較・効果'],
    ['論点・今後の進め方','決定事項の確認・スケジュール・担当'],
  ];
  items.forEach(([h,d],i)=>{ const y=1.75+i*1.0;
    circleNum(s,i+1,0.9,y+0.05);
    s.addText(h,{x:1.75,y,w:5,h:0.4,fontFace:F,fontSize:20,bold:true,color:C.navy,margin:0,isTextBox:true});
    s.addText(d,{x:1.75,y:y+0.42,w:10,h:0.35,fontFace:F,fontSize:14,color:C.ph,margin:0,isTextBox:true});
  });
  footer(s,n++);
}

// 3 目的・背景
{ const s = pres.addSlide(); title(s,'会議の目的・背景');
  card(s,0.6,1.6,5.9,4.9); card(s,6.85,1.6,5.9,4.9);
  const col=(x,h,lines)=>{ s.addText(h,{x:x+0.35,y:1.85,w:5.2,h:0.5,fontFace:F,fontSize:20,bold:true,color:C.navy,margin:0,isTextBox:true});
    s.addText(lines.map((t,i)=>({text:t,options:{bullet:true,color:C.ph,breakLine:i<lines.length-1}})),{x:x+0.35,y:2.5,w:5.2,h:3.7,fontFace:F,fontSize:15,valign:'top',paraSpaceAfter:10,margin:0,isTextBox:true}); };
  col(0.6,'背景',['【資材価格・労務単価の急激な変動状況】','【現行の計算事務で生じている課題】','【対象となる工事・契約の状況】']);
  col(6.85,'本日の目的',['【インフレスライドの考え方を共有する】','【計算方法の簡略化案について意見を伺う】','【今後の進め方を決定する】']);
  s.addText([{text:'本日決めたいこと：',options:{bold:true,color:C.navy}},{text:'【例：簡略化案の方向性（案A／案B）の決定】',options:{color:C.ph}}],{x:0.6,y:6.65,w:12.1,h:0.4,fontFace:F,fontSize:15,margin:0,isTextBox:true});
  footer(s,n++);
}

// 4 section
section(1,'インフレスライドとは','制度の概要・適用要件・手続きの流れ'); n++;

// 5 スライド条項の比較
{ const s = pres.addSlide(); title(s,'スライド条項の種類と比較','※公共工事標準請負契約約款 第26条（内容は要確認・適宜修正）');
  const hdr = (t)=>({text:t,options:{bold:true,color:C.white,fill:{color:C.navy},align:'center'}});
  const lab = (t)=>({text:t,options:{bold:true,color:C.navy,fill:{color:C.light}}});
  const rows = [
    [hdr(''),hdr('全体スライド'),hdr('単品スライド'),{text:'インフレスライド',options:{bold:true,color:C.white,fill:{color:C.amber},align:'center'}}],
    [lab('根拠条項'),'第26条第1〜4項','第26条第5項','第26条第6項'],
    [lab('対象'),'【比較的緩やかな価格水準の変動】','【特定の主要資材の価格変動】','【賃金水準・物価水準の急激な変動】'],
    [lab('適用条件'),'【契約締結から12か月経過 等】','【工期内で価格が著しく変動】','【工期内で急激な変動が発生】'],
    [lab('受注者負担'),'【残工事費の1.5%】','【対象工事費の1%】','【残工事費の1%】'],
    [lab('備考'),'【記入】','【記入】','【記入】'],
  ].map((r,ri)=> ri===0? r : r.map((c,ci)=> ci===0? c : (typeof c==='string'? {text:c,options:{color:c.startsWith('【')?C.ph:C.text}} : c)));
  s.addTable(rows,{x:0.6,y:1.8,w:12.1,colW:[2.2,3.3,3.3,3.3],rowH:[0.55,0.75,0.75,0.75,0.75,0.75],fontFace:F,fontSize:14,valign:'middle',border:{type:'solid',pt:0.75,color:C.line}});
  footer(s,n++);
}

// 6 手続きフロー
{ const s = pres.addSlide(); title(s,'適用の手続きフロー','【実際の運用に合わせてステップを修正してください】');
  const steps=['スライド\n請求','基準日の\n設定','残工事量の\n確認','変動額の\n算定','協議・\n変更契約'];
  const w=2.05,g=0.46,x0=0.6,y=2.2;
  steps.forEach((t,i)=>{ const x=x0+i*(w+g);
    card(s,x,y,w,1.5,i===3?C.amber:C.navy);
    s.addText(t,{x,y,w,h:1.5,fontFace:F,fontSize:17,bold:true,color:C.white,align:'center',valign:'middle',margin:0,isTextBox:true});
    if(i<steps.length-1) s.addShape(pres.shapes.RIGHT_TRIANGLE? pres.shapes.ISOSCELES_TRIANGLE:pres.shapes.ISOSCELES_TRIANGLE,{x:x+w+0.12,y:y+0.58,w:0.3,h:0.34,rotate:90,fill:{color:C.line},line:{color:C.line}});
    s.addText([{text:'【担当】',options:{bold:true,color:C.navy,breakLine:true}},{text:'【内容・留意点を記入】',options:{color:C.ph}}],{x,y:y+1.75,w,h:2.4,fontFace:F,fontSize:13,valign:'top',margin:0.05,isTextBox:true});
  });
  footer(s,n++);
}

// 7 section
section(2,'計算方法','計算式・算定手順・計算例'); n++;

// 8 計算式
{ const s = pres.addSlide(); title(s,'インフレスライド額の計算式');
  card(s,0.6,1.6,12.1,1.7,C.navy);
  s.addText([{text:'スライド額 S ＝ ',options:{color:C.white}},{text:'( P2 − P1 )',options:{color:C.amber}},{text:' − P1 × 1/100',options:{color:C.white}}],{x:0.6,y:1.6,w:12.1,h:1.7,fontFace:F,fontSize:34,bold:true,align:'center',valign:'middle',margin:0,isTextBox:true});
  const defs=[['S','インフレスライド額','【定義・留意点を記入】'],['P1','変動前残工事代金額','【例：基準日における残工事量を変動前の単価で積算した額 × 落札率】'],['P2','変動後残工事代金額','【例：基準日における残工事量を変動後の単価で積算した額 × 落札率】'],['1/100','受注者負担分','【残工事費の1%を受注者が負担】']];
  defs.forEach(([k,h,d],i)=>{ const y=3.7+i*0.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y,w:1.2,h:0.6,fill:{color:C.light},line:{color:C.light},rectRadius:0.08});
    s.addText(k,{x:0.6,y,w:1.2,h:0.6,fontFace:'Arial',fontSize:18,bold:true,color:C.navy,align:'center',valign:'middle',margin:0,isTextBox:true});
    s.addText(h,{x:2.05,y,w:3.2,h:0.6,fontFace:F,fontSize:16,bold:true,color:C.text,valign:'middle',margin:0,isTextBox:true});
    s.addText(d,{x:5.3,y,w:7.4,h:0.6,fontFace:F,fontSize:14,color:C.ph,valign:'middle',margin:0,isTextBox:true});
  });
  footer(s,n++);
}

// 9 計算手順
{ const s = pres.addSlide(); title(s,'計算の手順');
  const steps=[['基準日を決定する','【請求日を基準日とする等、運用ルールを記入】'],['残工事量を確定する','【出来高確認の方法・対象外とするものを記入】'],['変動前後の単価を設定する','【使用する単価（設計単価・物価資料等）を記入】'],['P1・P2を積算する','【積算方法・落札率の扱いを記入】'],['スライド額を算定する','【端数処理・請求要件の確認方法を記入】']];
  s.addShape(pres.shapes.LINE,{x:1.175,y:1.95,w:0,h:4.4,line:{color:C.line,width:2}});
  steps.forEach(([h,d],i)=>{ const y=1.7+i*0.95;
    circleNum(s,i+1,0.9,y,0.55,C.navy);
    s.addText(h,{x:1.8,y:y-0.02,w:4.5,h:0.6,fontFace:F,fontSize:18,bold:true,color:C.navy,valign:'middle',margin:0,isTextBox:true});
    s.addText(d,{x:6.3,y:y-0.02,w:6.4,h:0.6,fontFace:F,fontSize:14,color:C.ph,valign:'middle',margin:0,isTextBox:true});
  });
  footer(s,n++);
}

// 10 計算例
{ const s = pres.addSlide(); title(s,'計算例','【数値はダミーです。実際の事例に置き換えてください】');
  const hdr=(t)=>({text:t,options:{bold:true,color:C.white,fill:{color:C.navy},align:'center'}});
  const r=(a,b,c)=>[{text:a,options:{bold:true,color:C.navy,fill:{color:C.light}}},{text:b,options:{align:'right',color:C.ph}},{text:c,options:{color:C.ph}}];
  s.addTable([[hdr('項目'),hdr('金額（円）'),hdr('備考')],r('変動前残工事代金額 P1','【XXX,XXX,XXX】','【算定根拠】'),r('変動後残工事代金額 P2','【XXX,XXX,XXX】','【算定根拠】'),r('変動額 P2 − P1','【XX,XXX,XXX】',''),r('受注者負担 P1 × 1%','【X,XXX,XXX】',''),[{text:'スライド額 S',options:{bold:true,color:C.white,fill:{color:C.amber}}},{text:'【XX,XXX,XXX】',options:{bold:true,align:'right',color:C.text,fill:{color:'FBEBD6'}}},{text:'',options:{fill:{color:'FBEBD6'}}}]],
    {x:0.6,y:1.8,w:8.0,colW:[3.0,2.3,2.7],rowH:0.62,fontFace:F,fontSize:14,valign:'middle',border:{type:'solid',pt:0.75,color:C.line}});
  card(s,9.0,1.8,3.7,3.72,C.navy);
  s.addText('スライド額',{x:9.2,y:2.1,w:3.3,h:0.4,fontFace:F,fontSize:16,color:'B9C8D4',align:'center',margin:0,isTextBox:true});
  s.addText('【XX.X】',{x:9.2,y:2.7,w:3.3,h:1.1,fontFace:'Arial',fontSize:48,bold:true,color:C.amber,align:'center',margin:0,isTextBox:true});
  s.addText('百万円',{x:9.2,y:3.8,w:3.3,h:0.4,fontFace:F,fontSize:16,color:C.white,align:'center',margin:0,isTextBox:true});
  s.addText('請負代金額の【X.X】%',{x:9.2,y:4.5,w:3.3,h:0.4,fontFace:F,fontSize:14,color:'B9C8D4',align:'center',margin:0,isTextBox:true});
  s.addText('【補足・注意点を記入】',{x:0.6,y:6.0,w:12.1,h:0.5,fontFace:F,fontSize:14,color:C.ph,margin:0,isTextBox:true});
  footer(s,n++);
}

// 11 section
section(3,'計算方法の簡略化','現行の課題・簡略化案の比較・期待される効果'); n++;

// 12 現行の課題
{ const s = pres.addSlide(); title(s,'現行の計算方法の課題');
  const items=[['事務負担','【例：全単価の再積算に時間がかかる】'],['算定期間','【例：請求から変更契約まで長期化】'],['確認・審査','【例：根拠資料の確認範囲が広い】']];
  items.forEach(([h,d],i)=>{ const x=0.6+i*4.1, w=3.9;
    card(s,x,1.7,w,3.9);
    circleNum(s,'!',x+0.35,2.0,0.7,C.amber);
    s.addText(`課題${i+1}`,{x:x+0.35,y:2.95,w:w-0.7,h:0.35,fontFace:F,fontSize:13,color:C.ph,margin:0,isTextBox:true});
    s.addText(h,{x:x+0.35,y:3.3,w:w-0.7,h:0.5,fontFace:F,fontSize:20,bold:true,color:C.navy,margin:0,isTextBox:true});
    s.addText(d+'\n【詳細を記入】',{x:x+0.35,y:3.95,w:w-0.7,h:1.4,fontFace:F,fontSize:14,color:C.ph,valign:'top',margin:0,paraSpaceAfter:6,isTextBox:true});
  });
  footer(s,n++);
}

// 13 簡略化案の比較
{ const s = pres.addSlide(); title(s,'簡略化案の比較');
  const hdr=(t,fill=C.navy)=>({text:t,options:{bold:true,color:C.white,fill:{color:fill},align:'center'}});
  const lab=(t)=>({text:t,options:{bold:true,color:C.navy,fill:{color:C.light}}});
  const p=(t)=>({text:t,options:{color:C.ph}});
  const m=(t)=>({text:t,options:{color:C.ph,align:'center'}});
  s.addTable([
    [hdr(''),hdr('現行'),hdr('案A【名称】'),hdr('案B【名称】'),hdr('案C【名称】')],
    [lab('概要'),p('【全単価を再積算】'),p('【例：主要資材のみ再積算】'),p('【例：指数による一括補正】'),p('【記入】')],
    [lab('メリット'),p('【記入】'),p('【記入】'),p('【記入】'),p('【記入】')],
    [lab('デメリット'),p('【記入】'),p('【記入】'),p('【記入】'),p('【記入】')],
    [lab('算定精度'),m('◎'),m('【○】'),m('【△】'),m('【-】')],
    [lab('事務負担'),m('×'),m('【○】'),m('【◎】'),m('【-】')],
  ],{x:0.6,y:1.6,w:12.1,colW:[1.9,2.55,2.55,2.55,2.55],rowH:[0.55,0.95,0.95,0.95,0.6,0.6],fontFace:F,fontSize:13,valign:'middle',border:{type:'solid',pt:0.75,color:C.line}});
  s.addText([{text:'推奨案：',options:{bold:true,color:C.amber}},{text:'【案X】　理由：【記入】',options:{color:C.ph}}],{x:0.6,y:6.5,w:12.1,h:0.4,fontFace:F,fontSize:15,margin:0,isTextBox:true});
  footer(s,n++);
}

// 14 効果
{ const s = pres.addSlide(); title(s,'簡略化による効果（試算）');
  const stat=(x,label,val,unit,dark)=>{ card(s,x,1.8,3.7,3.4,dark?C.navy:C.light);
    s.addText(label,{x:x+0.3,y:2.05,w:3.1,h:0.4,fontFace:F,fontSize:16,bold:true,color:dark?'B9C8D4':C.navy,align:'center',margin:0,isTextBox:true});
    s.addText(val,{x:x+0.3,y:2.7,w:3.1,h:1.2,fontFace:'Arial',fontSize:54,bold:true,color:dark?C.amber:C.navy,align:'center',margin:0,isTextBox:true});
    s.addText(unit,{x:x+0.3,y:4.1,w:3.1,h:0.4,fontFace:F,fontSize:15,color:dark?C.white:C.ph,align:'center',margin:0,isTextBox:true}); };
  stat(0.6,'現行','【XX】','日／件（作業日数）',false);
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE,{x:4.37,y:3.3,w:0.4,h:0.4,rotate:90,fill:{color:C.line},line:{color:C.line}});
  stat(4.8,'簡略化後','【XX】','日／件（作業日数）',true);
  stat(9.0,'削減効果','【XX】','%（削減率）',false);
  s.addText([{text:'試算の前提：',options:{bold:true,color:C.navy}},{text:'【対象件数・前提条件を記入】',options:{color:C.ph,breakLine:true}},{text:'精度への影響：',options:{bold:true,color:C.navy}},{text:'【現行との差額の試算結果を記入】',options:{color:C.ph}}],{x:0.6,y:5.5,w:12.1,h:1.1,fontFace:F,fontSize:14,valign:'top',paraSpaceAfter:6,margin:0,isTextBox:true});
  footer(s,n++);
}

// 15 section
section(4,'論点・今後の進め方','本日の論点・スケジュール・担当'); n++;

// 16 論点
{ const s = pres.addSlide(); title(s,'本日ご議論いただきたい論点');
  const qs=['【論点1：例）簡略化案の方向性をどれにするか】','【論点2：例）適用対象（工事種別・規模）の範囲】','【論点3：例）受注者への説明・周知方法】'];
  qs.forEach((q,i)=>{ const y=1.7+i*1.55;
    card(s,0.6,y,12.1,1.3);
    s.addText('Q'+(i+1),{x:0.9,y,w:1.0,h:1.3,fontFace:'Arial',fontSize:32,bold:true,color:C.amber,valign:'middle',margin:0,isTextBox:true});
    s.addText([{text:q,options:{bold:true,color:C.navy,breakLine:true}},{text:'【補足・判断材料を記入】',options:{color:C.ph,fontSize:13}}],{x:2.0,y,w:10.4,h:1.3,fontFace:F,fontSize:17,valign:'middle',margin:0,isTextBox:true});
  });
  footer(s,n++);
}

// 17 スケジュール
{ const s = pres.addSlide(); title(s,'今後のスケジュール');
  const ms=[['【XX月】','方針決定'],['【XX月】','運用ルール案作成'],['【XX月】','関係者調整・周知'],['【XX月】','試行開始'],['【XX月】','本格運用']];
  const x0=1.7,x1=11.6,y=3.3; s.addShape(pres.shapes.LINE,{x:x0,y,w:x1-x0,h:0,line:{color:C.line,width:3}});
  ms.forEach(([m,t],i)=>{ const cx=x0+i*((x1-x0)/(ms.length-1));
    s.addShape(pres.shapes.OVAL,{x:cx-0.18,y:y-0.18,w:0.36,h:0.36,fill:{color:i===0?C.amber:C.navy},line:{color:C.white,width:2}});
    s.addText(m,{x:cx-1.1,y:y-0.9,w:2.2,h:0.45,fontFace:F,fontSize:16,bold:true,color:C.navy,align:'center',margin:0,isTextBox:true});
    s.addText(t,{x:cx-1.1,y:y+0.4,w:2.2,h:0.45,fontFace:F,fontSize:15,bold:true,color:C.text,align:'center',margin:0,isTextBox:true});
    s.addText('【詳細】',{x:cx-1.1,y:y+0.85,w:2.2,h:0.4,fontFace:F,fontSize:13,color:C.ph,align:'center',margin:0,isTextBox:true});
  });
  footer(s,n++);
}

// 18 まとめ / 次回アクション
{ const s = pres.addSlide(); title(s,'決定事項とネクストアクション');
  card(s,0.6,1.6,4.6,4.9,C.navy);
  s.addText('本日の決定事項',{x:0.9,y:1.85,w:4.0,h:0.5,fontFace:F,fontSize:20,bold:true,color:C.amber,margin:0,isTextBox:true});
  s.addText(['【決定事項1】','【決定事項2】','【決定事項3】'].map((t,i,a)=>({text:t,options:{bullet:true,breakLine:i<a.length-1}})),{x:0.9,y:2.55,w:4.0,h:3.6,fontFace:F,fontSize:15,color:'D5DEE5',valign:'top',paraSpaceAfter:10,margin:0,isTextBox:true});
  const hdr=(t)=>({text:t,options:{bold:true,color:C.white,fill:{color:C.navy},align:'center'}});
  const row=(a)=>[{text:a,options:{color:C.ph}},{text:'【担当】',options:{color:C.ph,align:'center'}},{text:'【XX/XX】',options:{color:C.ph,align:'center'}}];
  s.addTable([[hdr('アクション'),hdr('担当'),hdr('期限')],row('【アクション1】'),row('【アクション2】'),row('【アクション3】'),row('【アクション4】')],
    {x:5.6,y:1.6,w:7.1,colW:[4.1,1.5,1.5],rowH:0.6,fontFace:F,fontSize:14,valign:'middle',border:{type:'solid',pt:0.75,color:C.line}});
  s.addText([{text:'次回会議：',options:{bold:true,color:C.navy}},{text:'【XX月XX日（X）XX:XX〜】',options:{color:C.ph}}],{x:5.6,y:5.0,w:7.1,h:0.4,fontFace:F,fontSize:15,margin:0,isTextBox:true});
  footer(s,n++);
}

// 19 参考資料
{ const s = pres.addSlide(); title(s,'参考資料');
  const refs=['公共工事標準請負契約約款 第26条（賃金又は物価の変動に基づく請負代金額の変更）','【発注機関のインフレスライド運用マニュアル等】','【物価資料・単価表の出典】','【その他参考資料】'];
  refs.forEach((t,i)=>{ const y=1.7+i*0.9;
    circleNum(s,i+1,0.6,y,0.5,C.navy2);
    s.addText(t,{x:1.4,y,w:11.3,h:0.5,fontFace:F,fontSize:15,color:t.startsWith('【')?C.ph:C.text,valign:'middle',margin:0,isTextBox:true});
  });
  footer(s,n++);
}

pres.writeFile({ fileName: require('path').join(__dirname,'inflation_slide_template.pptx') }).then(f=>console.log(f));
