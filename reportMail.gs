// @ts-nocheck
function createDraft() {
   var ui = SpreadsheetApp.getUi()
  var al = ui.alert("確認：①タイトルの有無　②シートが自分の担任化　③差し込みの反映が適切か", ui.ButtonSet.YES_NO)
  switch(al){
    case ui.Button.YES:  
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sousintabu= sheet.getSheetByName('送信タブ');
  const lastRow = sousintabu.getRange(sousintabu.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  for(let i = 3; i <= lastRow; i++){
    
    // 送信タブ
    const fastName = sousintabu.getRange(i, 17).getValue(); //名
    const lastName = sousintabu.getRange(i, 16).getValue(); //姓
    const address = sousintabu.getRange(i, 14).getValue(); //生徒メール
//    const cc_mail = sousintabu.getRange(i, 15).getValue(); //保護者メール
    const bcc_mail = sousintabu.getRange(i, 22).getValue(); //bcc
    const shintyoku = sousintabu.getRange(i, 18).getValue(); //現在の進捗
    const shigatu = sousintabu.getRange(i, 5).getValue(); //4月以降のレポート進捗
    const gogatu = sousintabu.getRange(i, 6).getValue();
    const rokugatu = sousintabu.getRange(i, 7).getValue();
    const shitigatu = sousintabu.getRange(i, 8).getValue();
    const hatigatu = sousintabu.getRange(i, 9).getValue();
    const kugatu = sousintabu.getRange(i, 10).getValue();
    const jugatu = sousintabu.getRange(i, 11).getValue();
    const juichigatu = sousintabu.getRange(i, 12).getValue();
    const junigatu = sousintabu.getRange(i, 13).getValue();
    
    const shigatuNum = Math.floor(Number(shigatu) * 1000) /10;
    const gogatuNum = Math.floor(Number(gogatu) * 1000) / 10;
    const rokugatuNum = Math.floor(Number(rokugatu) * 1000) / 10;
    const shitigatuNum = Math.floor(Number(shitigatu) * 1000) / 10;
    const hatigatuNum = Math.floor(Number(hatigatu) * 1000) / 10;
    const kugatuNum = Math.floor(Number(kugatu) * 1000) / 10;
    const jugatuNum = Math.floor(Number(jugatu) * 1000) / 10;
    const juichigatuNum = Math.floor(Number(juichigatu) * 1000) / 10;
    const junigatuNum = Math.floor(Number(junigatu) * 1000) / 10;

    
    const bcc1 = sousintabu.getRange("V3").getValue();
    const bcc2 = sousintabu.getRange("V4").getValue();
    const bcc3 = sousintabu.getRange("V5").getValue();

    var bccMail = bcc1 + ',' + bcc2 + ',' + bcc3;
    
    const options = {
      bcc: bccMail //BCC
    };
    const subject = sousintabu.getRange(3, 19).getValue(); //タイトル
    //const mokuhyou = sousintabu.getRange(3, 12).getValue(); //今月の目標
    const tannin_name = sousintabu.getRange(3, 21).getValue(); //担任名
    
    //本文タブ
    const text = sheet.getSheetByName('本文').getRange(1, 1).getValue(); //本文
    const body = text
    .replace(/{名}/g,fastName)
    .replace(/{姓}/g,lastName)
//    .replace(/{今月の目標}/g,mokuhyou)
    .replace(/{担任名}/g, tannin_name)
//    .replace(/{現在の進捗率}/g, shintyoku)
    .replace(/{4月}/g, shigatuNum)
    .replace(/{5月}/g, gogatuNum)
    .replace(/{6月}/g, rokugatuNum)
    .replace(/{7月}/g, shitigatuNum)
    .replace(/{8月}/g, hatigatuNum)
    .replace(/{9月}/g, kugatuNum)
    .replace(/{10月}/g, jugatuNum)
    .replace(/{11月}/g, juichigatuNum)
    .replace(/{12月}/g, junigatuNum)

    GmailApp.createDraft(address, subject, body, options);
  //  console.log(bcc1);
  }
  }
}
