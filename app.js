/* Откуп на терен — логика на теренската апликација. Податоци: localStorage на уредот + SharePoint преку zaednicko.js */
(function(){
  "use strict";

  var VIDOVI = {kravjo:"Кравјо", ovco:"Овчо", kozjo:"Козјо"};
  var K_VNES = "otkup_vnes_v3", K_SPISOK = "otkup_spisok_v3", K_PODESI = "otkup_podesi_v3";
  var najaven = false;

  function el(id){ return document.getElementById(id); }
  function citaj(k, ako){ try{ var s = localStorage.getItem(k); return s ? JSON.parse(s) : ako; }catch(e){ return ako; } }
  function pisi(k, v){ try{ localStorage.setItem(k, JSON.stringify(v)); return true; }catch(e){ kazi("Меморијата на таблетот е полна — прати ги внесовите."); return false; } }
  function denes(){ var d = new Date(), p = function(n){ return n < 10 ? "0"+n : ""+n; };
    return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate()); }
  function kazi(t, ms){
    var d = document.createElement("div"); d.className = "poraka"; d.textContent = t;
    document.body.appendChild(d); setTimeout(function(){ d.remove(); }, ms || 2800);
  }
  function broj(s){ var n = parseFloat(String(s == null ? "" : s).replace(",", ".").trim()); return isNaN(n) ? null : n; }
  function fmt(n){ return String(Math.round(n * 10) / 10).replace(".", ","); }

  var VNES = citaj(K_VNES, []);
  var SPISOK = citaj(K_SPISOK, null);
  var PODESI = citaj(K_PODESI, {koj:"", reon:"", den:""});

  // ── екрани ───────────────────────────────────────────────
  var EKRANI = {Sken:["ekSken"], List:["ekList"], Prati:["ekPrati"]};
  var tekovenJazik = "Sken";
  function pokazi(koj){
    tekovenJazik = koj;
    ["Sken","List","Prati"].forEach(function(x){
      el("j"+x).setAttribute("aria-selected", x === koj ? "true" : "false");
    });
    var podeseno = !!(PODESI.koj && PODESI.reon);
    el("ekPodesi").hidden = !(koj === "Sken" && !podeseno);
    el("ekSken").hidden = !(koj === "Sken" && podeseno && el("ekVnes").hidden);
    el("ekList").hidden = koj !== "List";
    el("ekPrati").hidden = koj !== "Prati";
    if(koj !== "Sken") el("ekVnes").hidden = true;
    if(koj === "List") crtajSpisok();
    if(koj === "Prati"){ crtajStatusPrati(); crtajStatusSpisok(); }
  }
  el("jSken").addEventListener("click", function(){ pokazi("Sken"); });
  el("jList").addEventListener("click", function(){ pokazi("List"); });
  el("jPrati").addEventListener("click", function(){ pokazi("Prati"); });

  // ── врска ────────────────────────────────────────────────
  function crtajZnak(){
    var z = el("znak");
    if(!navigator.onLine){ z.textContent = "нема мрежа"; z.className = "znak lose"; return; }
    if(!najaven){ z.textContent = "не е најавен"; z.className = "znak lose"; return; }
    z.textContent = "поврзан"; z.className = "znak";
  }
  window.addEventListener("online", function(){ crtajZnak(); probajPrati(true); });
  window.addEventListener("offline", crtajZnak);

  // ── список на кооперанти ────────────────────────────────
  function vcitajSpisok(tivko){
    if(!navigator.onLine){ if(!tivko) kazi("Нема мрежа — пробај кога ќе има сигнал."); return Promise.resolve(); }
    if(!najaven){ if(!tivko) kazi("Најави се прво (јазиче Прати → Сметка)."); return Promise.resolve(); }
    return Otkup.kooperanti().then(function(lista){
      lista = lista.filter(function(x){ return x.aktiven; });
      if(!lista.length){ if(!tivko) kazi("Списокот во SharePoint е празен."); crtajStatusSpisok(); return; }
      SPISOK = {koga: new Date().toISOString(), lista: lista};
      pisi(K_SPISOK, SPISOK);
      napolniReoni(); crtajStatusSpisok();
      if(!tivko) kazi(lista.length + " кооперанти во списокот");
    }).catch(function(e){ if(!tivko) kazi(porakaGreska(e)); crtajStatusSpisok(); });
  }
  function porakaGreska(e){
    var m = (e && e.message) || "";
    if(m === "ne_najaven") return "Најави се прво.";
    if(m === "nema_mreza" || !navigator.onLine) return "Нема мрежа — пробај кога ќе има сигнал.";
    if(e && e.status === 403) return "Сметката нема пристап до SharePoint листата.";
    if(e && e.status === 404) return "Листата не постои — канцеларијата треба да ја направи (setup).";
    return "Не успеа: " + (m || "непозната грешка");
  }
  function najdi(sifra){
    if(!SPISOK || !SPISOK.lista) return null;
    sifra = String(sifra).trim();
    for(var i = 0; i < SPISOK.lista.length; i++) if(SPISOK.lista[i].sifra === sifra) return SPISOK.lista[i];
    return null;
  }
  function napolniReoni(){
    var s = el("reon"), sakan = s.value || PODESI.reon, vid = {};
    if(SPISOK && SPISOK.lista) SPISOK.lista.forEach(function(x){ if(x.reonBr) vid[x.reonBr] = x.reon || x.reonBr; });
    var kluci = Object.keys(vid).sort(function(a,b){ return Number(a) - Number(b); });
    s.innerHTML = "";
    if(!kluci.length){
      s.innerHTML = '<option value="">нема список — освежи на сигнал</option>'; return;
    }
    kluci.forEach(function(r){
      var o = document.createElement("option");
      o.value = r; o.textContent = "Реон " + vid[r];
      s.appendChild(o);
    });
    if(sakan && vid[sakan]) s.value = sakan;
  }
  function crtajStatusSpisok(){
    var h = el("statusSpisok");
    if(!SPISOK || !SPISOK.lista || !SPISOK.lista.length){
      h.innerHTML = '<div class="kazi lose">Нема список на телефонот. Освежи додека има сигнал.</div>'; return;
    }
    var d = new Date(SPISOK.koga);
    h.innerHTML = '<div style="font-size:19px;font-weight:700;font-family:\'JetBrains Mono\',monospace;">' +
      SPISOK.lista.length + '</div><div class="sitno" style="margin-top:2px;">кооперанти · освежено ' +
      d.toLocaleDateString("mk-MK") + " " + d.toLocaleTimeString("mk-MK", {hour:"2-digit", minute:"2-digit"}) + '</div>';
  }

  // ── QR од слика ─────────────────────────────────────────
  el("slikaj").addEventListener("click", function(){ el("foto").click(); });
  el("foto").addEventListener("change", function(){
    var f = this.files && this.files[0];
    this.value = "";
    if(!f) return;
    el("fotoPoraka").textContent = "Ја читам сликата…";
    var url = URL.createObjectURL(f);
    var img = new Image();
    img.onload = function(){
      var kod = procitajQR(img);
      URL.revokeObjectURL(url);
      if(kod){
        el("fotoPoraka").textContent = "Сликај го малиот квадрат до името на дневниот лист.";
        obrabotiKod(kod);
      } else {
        el("fotoPoraka").textContent = "Не го најдов кодот на сликата. Пробај поблиску и со подобро светло, или впиши ја шифрата подолу.";
      }
    };
    img.onerror = function(){ URL.revokeObjectURL(url); el("fotoPoraka").textContent = "Сликата не се отвора."; };
    img.src = url;
  });
  function procitajQR(img){
    if(typeof jsQR === "undefined") return null;
    // се проба на неколку големини: далечна слика бара помало намалување
    var probi = [1000, 700, 1400];
    for(var i = 0; i < probi.length; i++){
      var maks = probi[i];
      var w = img.naturalWidth, h = img.naturalHeight;
      var k = Math.min(1, maks / Math.max(w, h));
      var cv = document.createElement("canvas");
      cv.width = Math.round(w * k); cv.height = Math.round(h * k);
      var ctx = cv.getContext("2d", {willReadFrequently:true});
      ctx.drawImage(img, 0, 0, cv.width, cv.height);
      try{
        var d = ctx.getImageData(0, 0, cv.width, cv.height);
        var r = jsQR(d.data, d.width, d.height, {inversionAttempts:"attemptBoth"});
        if(r && r.data) return r.data;
      }catch(e){ /* следна големина */ }
    }
    return null;
  }

  // ── барање по шифра ─────────────────────────────────────
  el("rakaNajdi").addEventListener("click", function(){
    var s = el("rakaSifra").value.trim();
    if(s) obrabotiKod(s);
  });
  el("rakaSifra").addEventListener("keydown", function(e){
    if(e.key === "Enter"){ e.preventDefault(); el("rakaNajdi").click(); }
  });
  el("rakaSifra").addEventListener("input", function(){
    var q = this.value.trim().toLowerCase();
    var h = el("predlozi");
    if(q.length < 2 || !SPISOK || !SPISOK.lista){ h.innerHTML = ""; return; }
    var pogodoci = SPISOK.lista.filter(function(x){
      return x.sifra.toLowerCase().indexOf(q) === 0 || (x.ime || "").toLowerCase().indexOf(q) > -1;
    }).slice(0, 6);
    if(!pogodoci.length){ h.innerHTML = ""; return; }
    h.innerHTML = '<ul class="stavki" style="margin-top:10px;"></ul>';
    var ul = h.querySelector("ul");
    pogodoci.forEach(function(x){
      var li = document.createElement("li");
      li.innerHTML = '<span class="sifra">' + x.sifra + '</span><span class="lice">' + x.ime + '</span>';
      li.style.cursor = "pointer";
      li.addEventListener("click", function(){ el("rakaSifra").value = ""; h.innerHTML = ""; obrabotiKod(x.sifra); });
      ul.appendChild(li);
    });
  });

  // ── внес ────────────────────────────────────────────────
  var tekoven = null;
  function obrabotiKod(tekst){
    tekst = String(tekst).trim();
    var sifra = tekst, vid = "", i = tekst.indexOf("__");
    if(i > -1){ sifra = tekst.slice(0, i); vid = tekst.slice(i + 2); }
    var k = najdi(sifra);
    otvoriVnes(k || {sifra:sifra, ime:"", reon:"", reonBr:"", mesto:"", vidovi:[]}, vid, !k);
  }
  function otvoriVnes(koop, vid, nepoznat){
    var mozni = (koop.vidovi && koop.vidovi.length) ? koop.vidovi : ["kravjo","ovco","kozjo"];
    if(!vid && mozni.length === 1) vid = mozni[0];
    tekoven = {koop: koop, vid: vid || mozni[0]};
    var h = "";
    if(nepoznat){
      h += '<div class="kazi lose">Шифрата ' + koop.sifra + ' ја нема во списокот. Освежи го списокот, или провери го бројот.</div>';
    }
    h += '<div class="najden"><div class="ime">' + (koop.ime || "Непознат кооперант") + '</div>' +
         '<div class="det">' + koop.sifra + (koop.mesto ? " · " + koop.mesto : "") +
         (koop.reon ? " · реон " + koop.reon : "") + '</div>';
    if(vid) h += '<div class="vid">' + (VIDOVI[vid] || vid) + '</div>';
    h += '</div>';
    el("kutijaNajden").innerHTML = h;
    var trebaIzbor = !vid;
    el("poljeVid").hidden = !trebaIzbor;
    if(trebaIzbor){
      var s = el("vid");
      s.innerHTML = "";
      mozni.forEach(function(t){
        var o = document.createElement("option"); o.value = t; o.textContent = VIDOVI[t] || t; s.appendChild(o);
      });
    }
    el("litri").value = ""; el("temp").value = ""; el("zabel").value = "";
    el("ekSken").hidden = true; el("ekVnes").hidden = false;
    setTimeout(function(){ el("litri").focus(); }, 120);
  }
  el("nazad").addEventListener("click", function(){
    tekoven = null; el("ekVnes").hidden = true; el("ekSken").hidden = false;
  });

  el("zacuvaj").addEventListener("click", function(){
    if(!tekoven) return;
    var l = broj(el("litri").value);
    if(l === null || l < 0){ kazi("Впиши колку литри."); el("litri").focus(); return; }
    var vid = el("poljeVid").hidden ? tekoven.vid : el("vid").value;
    var reon = String(tekoven.koop.reonBr || PODESI.reon || "").trim();
    if(!/^[0-9]+$/.test(reon)){ kazi("Не знам во кој реон оди " + tekoven.koop.sifra + " — освежи го списокот."); return; }
    var zapis = {
      id: String(Date.now()) + "-" + Math.random().toString(36).slice(2, 7),
      den: PODESI.den, reon: reon, reonIme: tekoven.koop.reon || reon,
      sifra: tekoven.koop.sifra, ime: tekoven.koop.ime, vid: vid,
      litri: l, temp: el("temp").value.trim(), z: el("zabel").value.trim(),
      koj: PODESI.koj, vreme: new Date().toISOString(), pratem: null
    };
    VNES.push(zapis); pisi(K_VNES, VNES);
    kazi((tekoven.koop.ime || tekoven.koop.sifra) + ": " + fmt(l) + " л");
    tekoven = null; el("ekVnes").hidden = true; el("ekSken").hidden = false;
    crtajSpisok(); crtajStatusPrati(); osveziBrojka();
    probajPrati(true);
  });

  // ── список за денот ─────────────────────────────────────
  function crtajSpisok(){
    var host = el("spisok");
    var d = PODESI.den || denes();
    el("listNaslov").textContent = "Внесено за " + d;
    var moi = VNES.filter(function(v){ return v.den === d; });
    if(!moi.length){
      host.innerHTML = '<div class="nista"><strong>Уште ништо за овој ден</strong>' +
        'Сликај го QR-кодот од дневниот лист, па впиши ги литрите.</div>';
      return;
    }
    var vk = 0;
    var ul = document.createElement("ul");
    ul.className = "stavki";
    moi.slice().reverse().forEach(function(v){
      vk += Number(v.litri) || 0;
      var li = document.createElement("li");
      if(v.pratem) li.className = "gotova";
      li.innerHTML = '<span class="sifra">' + v.sifra + '</span>' +
        '<span class="lice">' + (v.ime || "—") + ' <small>' + (VIDOVI[v.vid] || v.vid) + '</small></span>' +
        '<span class="kolku">' + fmt(v.litri) + ' л</span>';
      if(v.pratem){
        var k = document.createElement("span"); k.className = "kvacka"; k.textContent = "✓"; k.title = "пратено";
        li.appendChild(k);
      } else {
        var b = document.createElement("button");
        b.className = "trgni"; b.type = "button"; b.textContent = "✕";
        b.setAttribute("aria-label", "Избриши го внесот за " + (v.ime || v.sifra));
        b.addEventListener("click", function(){ potvrdiBrisenje(v, li); });
        li.appendChild(b);
      }
      ul.appendChild(li);
    });
    host.innerHTML = "";
    host.appendChild(ul);
    var z = document.createElement("div");
    z.className = "zbir";
    z.innerHTML = "<span>" + moi.length + " записи</span><span class='broj'>" + fmt(vk) + " л</span>";
    host.appendChild(z);
  }
  // confirm() не работи во артефакт — потврдата се прави во самата страница
  function potvrdiBrisenje(v, li){
    if(li.querySelector(".potvrda")) return;
    var d = document.createElement("div");
    d.className = "potvrda";
    d.style.cssText = "flex:0 0 100%; display:flex; gap:8px; padding-top:9px;";
    d.innerHTML = '<button class="kop crven" style="padding:9px;font-size:14px;">Избриши</button>' +
                  '<button class="kop tivok" style="padding:9px;font-size:14px;margin-top:0;">Остави</button>';
    li.style.flexWrap = "wrap";
    li.appendChild(d);
    d.children[0].addEventListener("click", function(){
      VNES = VNES.filter(function(x){ return x.id !== v.id; });
      pisi(K_VNES, VNES); crtajSpisok(); crtajStatusPrati(); osveziBrojka();
      kazi("Внесот е избришан");
    });
    d.children[1].addEventListener("click", function(){ d.remove(); li.style.flexWrap = ""; });
  }

  // ── праќање ─────────────────────────────────────────────
  function cekaat(){ return VNES.filter(function(v){ return !v.pratem; }); }
  function osveziBrojka(){
    var n = cekaat().length, b = el("brCeka");
    b.textContent = n; b.hidden = !n;
  }
  function crtajStatusPrati(){
    var n = cekaat().length;
    el("statusPrati").innerHTML = n
      ? '<div style="font-size:30px;font-weight:700;font-family:\'JetBrains Mono\',monospace;line-height:1;">' + n + '</div>' +
        '<div class="sitno" style="margin-top:4px;">' + (n === 1 ? "внес чека" : "внеса чекаат") + ' да се пратат</div>'
      : '<div class="nista"><strong>Сè е пратено</strong>Канцеларијата ги има сите внесови.</div>';
    el("pratiKop").disabled = !n;
  }
  var pratam = false;
  function probajPrati(tivko){
    if(pratam || !najaven || !navigator.onLine) return Promise.resolve();
    var red = cekaat();
    if(!red.length) return Promise.resolve();
    pratam = true;
    var kop = el("pratiKop"), star = kop.textContent;
    var dobri = 0, losi = 0;
    function sledno(i){
      if(i >= red.length){
        pratam = false;
        kop.textContent = star;
        pisi(K_VNES, VNES);
        crtajSpisok(); crtajStatusPrati(); osveziBrojka();
        if(!tivko) kazi(losi ? ("Пратени " + dobri + ", не успеаја " + losi) : ("Пратени сите " + dobri), 3600);
        return Promise.resolve();
      }
      var v = red[i];
      if(!tivko) kop.textContent = "Праќам " + (i + 1) + " од " + red.length;
      return Otkup.pratiVnes(v)
        .then(function(){ v.pratem = new Date().toISOString(); dobri++; })
        .catch(function(e){ losi++; if(!tivko && i === 0) kazi(porakaGreska(e), 3600); })
        .then(function(){ return sledno(i + 1); });
    }
    return sledno(0);
  }
  el("pratiKop").addEventListener("click", function(){ probajPrati(false); });
  el("osvezi").addEventListener("click", function(){ vcitajSpisok(false); });

  // ── подесување ──────────────────────────────────────────
  el("pocni").addEventListener("click", function(){
    var koj = el("koj").value.trim();
    if(!koj){ kazi("Впиши го своето име."); el("koj").focus(); return; }
    if(!el("reon").value){ kazi("Нема список — освежи додека има сигнал."); return; }
    PODESI = {koj: koj, reon: el("reon").value, den: el("den").value || denes()};
    pisi(K_PODESI, PODESI);
    crtajPod(); pokazi("Sken"); crtajSpisok();
  });
  function crtajPod(){
    el("pod").textContent = (PODESI.koj && PODESI.reon)
      ? PODESI.koj + " · реон " + PODESI.reon + " · " + PODESI.den
      : "Ѓоргиеви ДООЕЛ";
  }

  // ── старт ───────────────────────────────────────────────
  el("den").value = PODESI.den || denes();
  el("koj").value = PODESI.koj || "";
  if(PODESI.den && PODESI.den !== denes()){ el("den").value = denes(); }
  napolniReoni(); crtajPod(); crtajSpisok(); crtajStatusPrati(); crtajStatusSpisok(); osveziBrojka(); crtajZnak();
  pokazi("Sken");

  function crtajSmetka(){
    var a = Otkup.smetka();
    najaven = !!a;
    el("smetkaIme").textContent = a ? (a.name || a.username) : "не е најавен";
    el("najava").textContent = a ? "Одјави се" : "Најави се";
    el("kartaNajavaPodesi").hidden = najaven;
    crtajZnak();
  }
  el("najava").addEventListener("click", function(){
    if(Otkup.smetka()){ if(cekaat().length){ kazi("Прво прати ги внесовите што чекаат.", 3600); return; } Otkup.odjavi(); }
    else Otkup.najavi().catch(function(e){ kazi(porakaGreska(e)); });
  });
  el("najava2").addEventListener("click", function(){ Otkup.najavi().catch(function(e){ kazi(porakaGreska(e)); }); });
  el("verzija").textContent = "Верзија " + (Otkup.config.verzija || "") + " · " + (Otkup.config.site || "");

  if("serviceWorker" in navigator){ navigator.serviceWorker.register("sw.js").catch(function(){}); }

  Otkup.init().then(function(a){
    crtajSmetka();
    if(a && !PODESI.koj && !el("koj").value && a.name) el("koj").value = a.name;
    if(najaven && navigator.onLine){
      var star = SPISOK && SPISOK.koga ? new Date(SPISOK.koga).getTime() : 0;
      var p = (Date.now() - star > 6*3600*1000) ? vcitajSpisok(true) : Promise.resolve();
      return p.then(function(){ return probajPrati(true); });
    }
  }).catch(function(e){ crtajSmetka(); if(!/ne_najaven|nema_mreza/.test(String(e && e.message))) kazi(porakaGreska(e)); });
})();
