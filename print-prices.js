/* Standalone print pricelist. Digital checkout and Apps Script stay unchanged. */
(()=>{'use strict';
const $=id=>document.getElementById(id),money=n=>'Rp '+n.toLocaleString('id-ID');
const parseQty=id=>{const n=Number($(id).value);return Number.isSafeInteger(n)&&n>0?n:0};
const report=(id,value,ok)=>{$(id).textContent=ok?money(value):'Periksa jumlah / minimum order';$('labelTotal')&&0;};
fetch('assets/print-prices.json').then(r=>{if(!r.ok)throw Error('Pricelist unavailable');return r.json()}).then(prices=>{
const fill=(id,values)=>values.forEach((v,i)=>{const opt=document.createElement('option');opt.value=i;opt.textContent=v;$(id).append(opt)});
fill('labelSize',prices.label.map(r=>r.size));fill('keySize',prices.keychains.map(r=>r.size));fill('a3Material',prices.a3.map(r=>r.name));
const state={};
const compute=()=>{
 const l=prices.label[Number($('labelSize').value)],lq=parseQty('labelQty'),lc=$('labelColor').value;
 $('labelMin').textContent='Minimal '+l.min.toLocaleString('id-ID')+' pcs • '+money(l[lc])+'/pcs';
 state.label=lq>=l.min?{total:lq*l[lc],desc:`Ukuran: ${l.size}\nWarna: ${lc==='one'?'1 warna':'2 warna'}\nJumlah: ${lq} pcs\nHarga/pcs: ${money(l[lc])}`} :null;
 report('labelTotal',state.label?.total||0,!!state.label);
 const wq=parseQty('wobblerQty');state.wobbler=wq>=15?{total:wq*4500,desc:`1 sisi, Art Paper 260 + mika, max 15 cm\nJumlah: ${wq} pcs\nHarga/pcs: Rp 4.500`}:null;
 report('wobblerTotal',state.wobbler?.total||0,!!state.wobbler);
 const k=prices.keychains[Number($('keySize').value)],ks=Number($('keySide').value),kq=parseQty('keyQty');state.keychain=kq>=15?{total:k.prices[ks]*kq,desc:`Ukuran: ${k.size}\nCetak: ${ks+1} sisi\nJumlah: ${kq} pcs\nHarga/pcs: ${money(k.prices[ks])}`} :null;
 report('keyTotal',state.keychain?.total||0,!!state.keychain);
 const aq=parseQty('a3Qty'),m=prices.a3[Number($('a3Material').value)],tier=aq<=50?0:aq<=100?1:aq<=500?2:3;
 const cut=$('a3Cut').checked,lam=$('a3Lam').checked;
 // Pricelist shows cutting and lamination by number of sheets, treated as order-level charges.
 const cutting=aq<=10?1000:aq<=30?1200:aq<=50?1700:aq<=99?2000:aq>100?2500:null;
 const laminating=aq<=50?6000:aq<=100?5000:aq>501?4000:null;
 const ok=aq>0&&(!cut||cutting!==null)&&(!lam||laminating!==null);
 state.a3=ok?{total:aq*m.prices[tier]+(cut?cutting:0)+(lam?laminating:0),desc:`Bahan: ${m.name}\nJumlah: ${aq} lembar\nHarga/lembar: ${money(m.prices[tier])}\nPotong: ${cut?money(cutting):'Tidak'}\nLaminasi: ${lam?money(laminating):'Tidak'}`} :null;
 report('a3Total',state.a3?.total||0,!!state.a3);
 const help=$('a3Total');help.title=ok?'Estimasi dari pricelist':'Rentang jumlah untuk biaya tambahan belum tercantum; hubungi admin';
 for(const type of ['label','wobbler','keychain','a3']){document.querySelector(`[data-order="${type}"]`).disabled=!state[type]}
};
['labelSize','labelColor','labelQty','wobblerQty','keySize','keySide','keyQty','a3Material','a3Qty','a3Cut','a3Lam'].forEach(id=>$(id).addEventListener('input',compute));
document.querySelectorAll('[data-order]').forEach(btn=>btn.addEventListener('click',()=>{const type=btn.dataset.order,st=state[type];if(!st)return;const names={label:'Label Satin Lokal',wobbler:'Wobbler Display Rak',keychain:'Gantungan Kunci Akrilik',a3:'Cetak A3+'};const msg=`Halo CETAKIN! Saya ingin memesan ${names[type]}.\n\n${st.desc}\nEstimasi total: ${money(st.total)}\n\nMohon konfirmasi desain, harga final, dan pembayaran.`;window.open('https://wa.me/6285196071937?text='+encodeURIComponent(msg),'_blank','noopener,noreferrer')}));
compute();
}).catch(()=>{document.querySelectorAll('.print-order').forEach(b=>b.disabled=true);document.querySelectorAll('.print-total').forEach(el=>el.textContent='Data harga gagal dimuat. Silakan hubungi admin.')});
})();
