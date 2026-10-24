/** CETAKIN backend — Google Apps Script bound to the target Google Sheet. */
const ORDERS='Pesanan', SUMMARY='Ringkasan';
// Setelah upload Logo-CETAKIN-Transparan.png ke Google Drive, masukkan ID file di sini.
const CETAKIN_LOGO_FILE_ID = '1DbPjSD9_S_f_QNjlxTOhoqyhGyI0VvF3';
function onOpen(){SpreadsheetApp.getUi().createMenu('CETAKIN').addItem('Tambah order cetak/custom','tambahOrderManual').addItem('Buat ulang nota baris terpilih','buatUlangNotaTerpilih').addToUi();}

const HEADERS=['ID Order','Tanggal','Nama','WhatsApp','Detail Produk','Total (Rp)','Status','Nota PDF','Terakhir Diubah','Metode Bayar','ID Transaksi ShopeePay','Tanggal Lunas'];
const PRICE={
 canva:{'Canva Pro':3000,'Canva Designer':4000},capcut:{'Sharing 1 bulan':20000,'Private 1 bulan':30000,'Private 7 hari':10000,'Sharing 7 hari':5000},
 chatgpt:{'Go 7 hari':15000,'Go 30 hari':30000,'Plus 7 hari':18000,'Plus 30 hari':35000},
 netflix:{'Sharing 1 user • 1 bulan':23000,'Sharing 2 user • 1 bulan':18000,'Semi Private • 1 bulan':30000,'Device TV • Sharing 1 user • 1 bulan':17000,'Device TV • Sharing 2 user • 1 bulan':13000,'Device TV • Semi Private • 1 bulan':20000},
 iqiyi:{'Sharing 1 bulan':12000,'Sharing 3 bulan':18000,'Sharing 1 tahun':22000,'Anti Limit 1 bulan':20000},
 spotify:{'Family Plan 1 bulan':18000,'Family Plan 2 bulan':32000,'Independent Plan 1 bulan':29000},
 primevideo:{"Sharing Anti Limit 1 bulan":8000,"Private 1 bulan":15000},
 wetv:{"Sharing 1 bulan":12000,"Private 1 bulan":27000},
 viu:{"Anti Limit 3 bulan":11000,"Anti Limit 6 bulan":21000},
 loklok:{"Standard Sharing 1 bulan":20500,"Basic Sharing 1 bulan":19500},
 bsstation:{"Sharing 1 bulan":9000,"Sharing 3 bulan":11000},
 disneyplus:{"Sharing 7 hari":10000,"Sharing 1 bulan • 6 user":25000,"Sharing 1 bulan Anti Limit • 3 user":35000},
 youtube:{"Family Plan 1 bulan":10000,"Independent Plan 1 bulan":15000},
 youtubemusic:{"Family Plan 1 bulan":10000,"Independent Plan 1 bulan":15000},
 vidio:{"TV Only 1 bulan":10000,"Sharing Mobile Only 1 bulan":19000,"Sharing All Device 1 bulan":35000,"Private Mobile Only 1 bulan":29000,"Private All Device 1 bulan":41000},
 applemusic:{"Family Plan 1 bulan":12000,"Independent Plan 1 bulan":15000},
 zoom:{"Private 7 hari":7000,"Private 14 hari":9000,"Private 1 bulan":13000},
 claude:{"Sharing 5 User 7 hari":45000,"Sharing 5 User 1 bulan":100000,"Private 1 bulan":335000}
};
const NAMES={canva:'Canva',capcut:'CapCut',chatgpt:'ChatGPT',netflix:'Netflix',iqiyi:'iQIYI Premium',spotify:'Spotify',primevideo:"Prime Video",wetv:"WeTV",viu:"Viu",loklok:"Loklok",bsstation:"BS Station",disneyplus:"Disney+",youtube:"YouTube",youtubemusic:"YouTube Music",vidio:"Vidio Platinum",applemusic:"Apple Music",zoom:"Zoom Pro",claude:"Claude AI"};
function setupCetakin(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(ORDERS)||ss.insertSheet(ORDERS);
 if(sh.getLastRow()===0)sh.appendRow(HEADERS);else sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
 sh.setFrozenRows(1);sh.getRange('A1:L1').setBackground('#1c1c1c').setFontColor('#ffffff').setFontWeight('bold');
 sh.getRange('B2:B').setNumberFormat('dd mmm yyyy hh:mm');sh.getRange('F2:F').setNumberFormat('"Rp" #,##0');sh.getRange('L2:L').setNumberFormat('dd mmm yyyy hh:mm');
 const rule=SpreadsheetApp.newDataValidation().requireValueInList(['MENUNGGU','LUNAS','BATAL'],true).build();sh.getRange('G2:G').setDataValidation(rule);
 let dash=ss.getSheetByName(SUMMARY)||ss.insertSheet(SUMMARY);dash.clear();
 const data=[['CETAKIN — LAPORAN PENJUALAN','Nilai'],['Total Omzet Lunas',`=SUMIF('${ORDERS}'!G2:G,"LUNAS",'${ORDERS}'!F2:F)`],['Total Pesanan',`=COUNTA('${ORDERS}'!A2:A)`],['Pesanan Lunas',`=COUNTIF('${ORDERS}'!G2:G,"LUNAS")`],['Menunggu Pembayaran',`=COUNTIF('${ORDERS}'!G2:G,"MENUNGGU")`],['Pesanan Batal',`=COUNTIF('${ORDERS}'!G2:G,"BATAL")`],['Omzet Bulan Ini',`=SUMIFS('${ORDERS}'!F2:F,'${ORDERS}'!G2:G,"LUNAS",'${ORDERS}'!L2:L,">="&EOMONTH(TODAY(),-1)+1,'${ORDERS}'!L2:L,"<"&EOMONTH(TODAY(),0)+1)`]];
 dash.getRange(1,1,data.length,2).setValues(data);dash.getRange('A1:B1').setBackground('#1c1c1c').setFontColor('white').setFontWeight('bold');dash.setColumnWidth(1,250);dash.setColumnWidth(2,210);dash.getRange('B2').setNumberFormat('"Rp" #,##0');dash.getRange('B7').setNumberFormat('"Rp" #,##0');
 const installed=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='onOrderStatusEdit');
 if(!installed)ScriptApp.newTrigger('onOrderStatusEdit').forSpreadsheet(ss).onEdit().create();
 sh.setColumnWidth(1,210); sh.setColumnWidth(3,190);sh.setColumnWidth(5,340);sh.setColumnWidth(8,290);sh.setColumnWidth(11,220);
 SpreadsheetApp.getUi().alert('Siap! Deploy sebagai Web app (Execute as me, akses Anyone) dan pasang URL /exec di script.js.');
}
function doPost(e){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  if(!e||!e.postData)throw Error('Request kosong');
  const x=JSON.parse(e.postData.contents),id=String(x.id||''),name=String(x.name||'').trim(),phone=String(x.phone||'').trim();
  if(!/^CTK-\d{8}-[A-Z0-9]{7,}$/.test(id)||name.length<2||name.length>80||!/^[+0-9 ()-]{8,25}$/.test(phone)||!Array.isArray(x.items)||x.items.length<1||x.items.length>30)throw Error('Data tidak valid');
  let total=0;const lines=[];
  for(const item of x.items){const opts=PRICE[item.id],price=opts&&opts[item.option];const qty=Number(item.quantity);
   if(!price||!Number.isInteger(qty)||qty<1||qty>99)throw Error('Produk atau jumlah tidak valid');total+=price*qty;lines.push(`${NAMES[item.id]} / ${item.option} × ${qty} = Rp${price*qty}`);}
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS);if(!sh)throw Error('Jalankan setupCetakin dahulu');
  const ids=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat():[];
  if(ids.includes(id))return ContentService.createTextOutput('DUPLICATE');
  // Escape formula injections into Sheet cells.
  const clean=t=> /^[=+@-]/.test(t)?"'"+t:t;
  sh.appendRow([id,new Date(),clean(name),clean(phone),clean(lines.join('\n')),total,'MENUNGGU','',new Date(),'ShopeePay QRIS','','']);
  return ContentService.createTextOutput('OK');
 }catch(err){console.error(err);return ContentService.createTextOutput('ERROR');}
 finally{lock.releaseLock();}
}
function onOrderStatusEdit(e){
 if(!e||!e.range||e.range.getSheet().getName()!==ORDERS||e.range.getColumn()!==7||e.range.getRow()<2||e.value!=='LUNAS')return;
 const sh=e.range.getSheet(),row=e.range.getRow();sh.getRange(row,9).setValue(new Date());
 if(!sh.getRange(row,12).getValue())sh.getRange(row,12).setValue(new Date());
 if(sh.getRange(row,8).getValue())return;
 try{const vals=sh.getRange(row,1,1,8).getValues()[0];const pdf=createInvoice_(vals,row);sh.getRange(row,8).setValue(pdf.getUrl());}
 catch(err){console.error(err);sh.getRange(row,8).setNote('Gagal membuat nota: '+err.message);}
}
function createInvoice_(r,row){
 const [id,date,name,phone,items,total]=r;
 const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS);
 const method=sheet.getRange(row,10).getDisplayValue()||'ShopeePay QRIS';
 const ref=sheet.getRange(row,11).getDisplayValue()||'Diverifikasi admin';
 const paidAt=sheet.getRange(row,12).getValue() || new Date();
 const formattedDate=Utilities.formatDate(new Date(paidAt),'Asia/Jakarta','dd MMMM yyyy, HH:mm')+' WIB';
 const money=v=>'Rp'+Number(v).toLocaleString('id-ID');
 const navy='#171B26',muted='#647083',blue='#039BE5',pink='#EC087B',yellow='#FFBD21',line='#DFE4E9';
 const doc=DocumentApp.create('NOTA CETAKIN '+id), body=doc.getBody();
 // Ukuran A5 portrait dalam satuan point (lebih ringkas untuk dibagikan via WhatsApp).
 body.setPageWidth(420).setPageHeight(595);
 body.setMarginTop(26).setMarginBottom(24).setMarginLeft(32).setMarginRight(32);
 const txt=(value,size,color,bold)=>{
   const p=body.appendParagraph(String(value));p.setSpacingBefore(0).setSpacingAfter(5);
   const t=p.editAsText();t.setFontFamily('Arial').setFontSize(size||9).setForegroundColor(color||navy);
   if(bold)t.setBold(true);return p;
 };
 const cell=(c,value,style={})=>{
   c.clear();const p=c.appendParagraph(String(value));p.setSpacingBefore(1).setSpacingAfter(1);
   const t=p.editAsText();t.setFontFamily('Arial').setFontSize(style.size||9).setForegroundColor(style.color||navy);
   if(style.bold)t.setBold(true);
   if(style.align)p.setAlignment(style.align);
   if(style.bg)c.setBackgroundColor(style.bg);
   return p;
 };
 if (!CETAKIN_LOGO_FILE_ID || CETAKIN_LOGO_FILE_ID==='ISI_ID_FILE_LOGO_DI_SINI')throw Error('ID logo belum diisi');
 const logoBlob=DriveApp.getFileById(CETAKIN_LOGO_FILE_ID).getBlob();
 const p=body.appendParagraph('');p.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(5);
 const logo=p.appendInlineImage(logoBlob),width=250;
 const ratio=logo.getHeight()/logo.getWidth();logo.setWidth(width).setHeight(Math.round(width*ratio));
 const title=txt('INVOICE  /  NOTA PEMBAYARAN',14,navy,true);title.setSpacingBefore(3);
 const subtitle=txt('Bukti pembayaran resmi • Status: LUNAS',9,blue,true);subtitle.setSpacingAfter(8);
 // Strip warna identitas CETAKIN.
 const strip=body.appendTable([['','','']]);strip.setBorderWidth(0);
 for(let i=0;i<3;i++){let c=strip.getCell(0,i);c.setBackgroundColor([blue,pink,yellow][i]);cell(c,' ',{size:3,bg:[blue,pink,yellow][i]});}
 body.appendParagraph('').setSpacingAfter(2);
 const info=body.appendTable([['NO. ORDER','TANGGAL PEMBAYARAN'],[id,formattedDate],['PELANGGAN','METODE BAYAR'],[String(name),String(method)]]);
 info.setBorderWidth(0);
 for(let ri=0;ri<4;ri++)for(let ci=0;ci<2;ci++)cell(info.getCell(ri,ci),info.getCell(ri,ci).getText(),{size:ri%2===0?7:9,bold:ri%2===1,color:ri%2===0?muted:navy});
 txt('WhatsApp pembeli: '+phone,8,muted);
 txt('Referensi pembayaran: '+ref,8,muted);
 txt('DETAIL PESANAN',10,navy,true).setSpacingBefore(9);
 const rows=[['DESKRIPSI','JUMLAH','NOMINAL']];
 const lines=String(items).split('\n').filter(x=>x.trim());
 for(const item of lines){
   // Detail checkout biasanya: Nama / Varian × 1 = Rp3000
   const m=item.match(/^(.*?)\s*[×x]\s*(\d+)\s*=\s*Rp\s*([\d.,]+)\s*$/i);
   if(m)rows.push([m[1].trim(),m[2],money(Number(m[3].replace(/[^0-9]/g,'')))]);
   else rows.push([item,'—','—']);
 }
 const table=body.appendTable(rows);table.setBorderColor(line).setBorderWidth(0.5);
 for(let ri=0;ri<rows.length;ri++)for(let ci=0;ci<3;ci++){
   const c=table.getCell(ri,ci);
   cell(c,rows[ri][ci],{size:ri===0?7:8,bold:ri===0,color:ri===0?'#FFFFFF':navy,bg:ri===0?navy:ri%2===0?'#F6F8FA':'#FFFFFF',align:ci===0?DocumentApp.HorizontalAlignment.LEFT:DocumentApp.HorizontalAlignment.RIGHT});
 }
 body.appendParagraph('').setSpacingAfter(2);
 const sum=body.appendTable([['TOTAL DIBAYAR',money(total)]]);sum.setBorderWidth(0);
 cell(sum.getCell(0,0),'TOTAL DIBAYAR',{bold:true,size:10,bg:'#EAF7FD'});
 cell(sum.getCell(0,1),money(total),{bold:true,size:14,bg:'#EAF7FD',align:DocumentApp.HorizontalAlignment.RIGHT});
 txt('✓ LUNAS — Pembayaran diverifikasi oleh admin CETAKIN',8,'#16834B',true).setSpacingBefore(7);
 body.appendHorizontalRule();
 const footer=txt('CETAKIN  •  PRINT / MERCHANDISE / CUSTOM EVERYTHING',7,navy,true);
 footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
 const address=txt('Jl. Jagal Abilowo No. 06, Kota Batu, Jawa Timur',7,muted);address.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
 const wa=txt('WhatsApp: +62 851-9607-1937  •  cetakin.store',7,muted);wa.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
 const thanks=txt('Terima kasih telah berbelanja di CETAKIN.',8,navy,true);thanks.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
 doc.saveAndClose();
 const folderName='CETAKIN NOTA PDF',iter=DriveApp.getFoldersByName(folderName);
 const folder=iter.hasNext()?iter.next():DriveApp.createFolder(folderName);
 const file=DriveApp.getFileById(doc.getId());
 try{return folder.createFile(file.getAs(MimeType.PDF).setName('NOTA-'+id+'.pdf'));}
 finally{file.setTrashed(true);}
}

/** Order percetakan dengan harga custom. Hanya admin pemilik sheet yang menggunakan menu ini. */
function tambahOrderManual(){
 const ui=SpreadsheetApp.getUi(),ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(ORDERS);
 if(!sh){ui.alert('Jalankan setupCetakin terlebih dahulu.');return;}
 const ask=(title,description)=>{const r=ui.prompt(title,description,ui.ButtonSet.OK_CANCEL);return r.getSelectedButton()===ui.Button.OK?r.getResponseText().trim():null;};
 const name=ask('Nama pelanggan','Nama pemesan (minimal 2 karakter)');if(name===null)return;
 const phone=ask('WhatsApp pelanggan','Nomor WhatsApp pemesan');if(phone===null)return;
 const product=ask('Deskripsi pesanan','Contoh: Banner 2×1 m, 1 pcs');if(product===null)return;
 const raw=ask('Total harga (Rupiah)','Masukkan angka saja, contoh 85000');if(raw===null)return;
 const price=Number(raw.replace(/[^0-9]/g,''));
 if(name.length<2||phone.length<8||!product||!Number.isSafeInteger(price)||price<=0){ui.alert('Data tidak valid. Pesanan tidak disimpan.');return;}
 const id='CTK-'+Utilities.formatDate(new Date(),'Asia/Jakarta','yyyyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();
 const safe=x=>/^[=+@-]/.test(x)?"'"+x:x;
 sh.appendRow([id,new Date(),safe(name),safe(phone),safe(product),price,'MENUNGGU','',new Date(),'ShopeePay QRIS','','']);
 ui.alert('Pesanan tersimpan: '+id+'\nPeriksa pembayaran ShopeePay dulu, lalu ubah status baris ini menjadi LUNAS.');
}
/** Perbaikan jika nota PDF gagal dibuat, untuk baris terpilih dengan status LUNAS. */
function buatUlangNotaTerpilih(){
 const ui=SpreadsheetApp.getUi(),sh=SpreadsheetApp.getActiveSheet(),row=sh.getActiveRange().getRow();
 if(sh.getName()!==ORDERS||row<2){ui.alert('Pilih baris pesanan di tab Pesanan.');return;}
 if(sh.getRange(row,7).getValue()!=='LUNAS'){ui.alert('Status harus LUNAS dan pembayaran sudah diverifikasi.');return;}
 if(sh.getRange(row,8).getValue()){ui.alert('Nota sudah ada di kolom H.');return;}
 try{const pdf=createInvoice_(sh.getRange(row,1,1,8).getValues()[0],row);sh.getRange(row,8).setValue(pdf.getUrl());ui.alert('Nota berhasil dibuat. Link tersedia di kolom H.');}
 catch(err){ui.alert('Gagal membuat nota: '+err.message);}
}

/** Pratinjau tampilan memakai data simulasi tanpa mengubah spreadsheet. */
function previewNotaPremium(){
 const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS);
 if(!sh||sh.getLastRow()<2)throw Error('Belum ada order untuk pratinjau.');
 const row=2,values=sh.getRange(row,1,1,8).getValues()[0];
 const file=createInvoice_(values,row);
 SpreadsheetApp.getUi().alert('Pratinjau nota siap. Buka link ini: '+file.getUrl());
}
