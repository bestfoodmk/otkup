/* Заедничко: најава со Microsoft (MSAL) и пристап до SharePoint листите преку Microsoft Graph.
   Го користат index.html (терен), setup.html (поставување) и kancelarija.html (канцеларија). */
(function(){
  "use strict";
  var C = window.OTKUP_CONFIG || {};
  var SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All", "https://graph.microsoft.com/User.Read"];
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var msalApp = null, spremno = null;

  function redirectUri(){ return location.origin + location.pathname; }

  function init(){
    if(spremno) return spremno;
    if(!window.msal){ spremno = Promise.reject(new Error("msal-browser не е вчитан")); return spremno; }
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: "https://login.microsoftonline.com/" + C.tenantId, redirectUri: redirectUri(), navigateToLoginRequestUrl: false },
      cache: { cacheLocation: "localStorage" },
      system: { allowRedirectInIframe: false }
    });
    spremno = msalApp.initialize().then(function(){ return msalApp.handleRedirectPromise(); }).then(function(r){
      if(r && r.account) msalApp.setActiveAccount(r.account);
      if(!msalApp.getActiveAccount()){ var a = msalApp.getAllAccounts(); if(a.length) msalApp.setActiveAccount(a[0]); }
      return smetka();
    });
    return spremno;
  }
  function smetka(){ return msalApp && msalApp.getActiveAccount ? msalApp.getActiveAccount() : null; }
  function najavi(){ return init().then(function(){ return msalApp.loginRedirect({ scopes: SCOPES, prompt: "select_account" }); }); }
  function odjavi(){ return init().then(function(){ return msalApp.logoutRedirect({ account: smetka(), postLogoutRedirectUri: redirectUri() }); }); }

  function token(){
    return init().then(function(){
      var a = smetka();
      if(!a) throw new Error("ne_najaven");
      return msalApp.acquireTokenSilent({ scopes: SCOPES, account: a }).then(function(r){ return r.accessToken; })
        .catch(function(e){
          if(!navigator.onLine) throw new Error("nema_mreza");
          if(e && e.name === "InteractionRequiredAuthError"){ return msalApp.acquireTokenRedirect({ scopes: SCOPES, account: a }); }
          throw e;
        });
    });
  }

  function graph(pateka, opcii){
    opcii = opcii || {};
    return token().then(function(t){
      var h = Object.assign({ "Authorization": "Bearer " + t, "Accept": "application/json" }, opcii.headers || {});
      if(opcii.body && typeof opcii.body !== "string"){ opcii.body = JSON.stringify(opcii.body); h["Content-Type"] = "application/json"; }
      var url = pateka.indexOf("https://") === 0 ? pateka : GRAPH + pateka;
      return fetch(url, { method: opcii.method || "GET", headers: h, body: opcii.body }).then(function(r){
        if(r.status === 204) return null;
        return r.text().then(function(txt){
          var j = null; try{ j = txt ? JSON.parse(txt) : null; }catch(e){}
          if(!r.ok){ var err = new Error((j && j.error && j.error.message) || ("Graph " + r.status)); err.status = r.status; err.graph = j; throw err; }
          return j;
        });
      });
    });
  }
  // Сите страници од една колекција (следи @odata.nextLink)
  function site(pateka, opcii){
    var se = [];
    function cekor(p){ return graph(p, opcii).then(function(j){ se = se.concat(j.value || []); return j["@odata.nextLink"] ? cekor(j["@odata.nextLink"]) : se; }); }
    return cekor(pateka);
  }

  var siteIdKes = null;
  function siteId(){
    if(siteIdKes) return Promise.resolve(siteIdKes);
    var k = "otkup_siteid_" + C.site; try{ var z = localStorage.getItem(k); if(z){ siteIdKes = z; return Promise.resolve(z); } }catch(e){}
    return graph("/sites/" + C.site).then(function(j){ siteIdKes = j.id; try{ localStorage.setItem(k, j.id); }catch(e){} return j.id; });
  }
  function lista(ime){ return siteId().then(function(id){ return "/sites/" + id + "/lists/" + encodeURIComponent(ime); }); }

  // Кооперанти: еден ред во листата = една шифра (едно лице + еден вид млеко)
  function kooperanti(){
    return lista(C.listaKooperanti).then(function(l){
      return site(l + "/items?$top=999&$expand=fields($select=Title,Sifra,Vid,Reon,ReonBr,Mesto,Otk,Aktiven,Prethodna)");
    }).then(function(redovi){
      return redovi.map(function(r){ var f = r.fields || {}; return {
        id: r.id, ime: f.Title || "", sifra: String(f.Sifra || "").trim(), vidovi: f.Vid ? [f.Vid] : [],
        reon: f.Reon || "", reonBr: f.ReonBr != null ? String(f.ReonBr) : "", mesto: f.Mesto || "", otk: f.Otk || "",
        aktiven: f.Aktiven !== false, prethodna: f.Prethodna || ""
      }; }).filter(function(x){ return x.sifra; });
    });
  }
  function dodajKooperant(x){
    return lista(C.listaKooperanti).then(function(l){ return graph(l + "/items", { method: "POST", body: { fields: {
      Title: x.ime, Sifra: x.sifra, Vid: (x.vidovi && x.vidovi[0]) || x.vid || "", Reon: x.reon || "", ReonBr: x.reonBr ? Number(x.reonBr) : null,
      Mesto: x.mesto || "", Otk: x.otk || "", Aktiven: x.aktiven !== false, Prethodna: x.prethodna || ""
    } } }); });
  }

  // Внесови од терен
  function postoiVnes(lokalenId){
    return lista(C.listaVnesovi).then(function(l){
      return graph(l + "/items?$expand=fields($select=LokalenId)&$filter=fields/LokalenId eq '" + lokalenId.replace(/'/g, "''") + "'",
        { headers: { "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly" } });
    }).then(function(j){ return !!(j.value && j.value.length); }).catch(function(){ return false; });
  }
  function pratiVnes(v){
    return postoiVnes(v.id).then(function(ima){
      if(ima) return { veke: true };
      return lista(C.listaVnesovi).then(function(l){ return graph(l + "/items", { method: "POST", body: { fields: {
        Title: v.ime || v.sifra, Sifra: v.sifra, Vid: v.vid, Litri: Number(v.litri), Temp: v.temp || "", Zabeleska: v.z || "",
        Den: v.den, Reon: v.reon, ReonIme: v.reonIme || "", Koj: v.koj, Vreme: v.vreme, LokalenId: v.id, Status: "Примено"
      } } }); });
    });
  }
  function vnesovi(den){
    return lista(C.listaVnesovi).then(function(l){
      var q = "/items?$top=999&$expand=fields($select=Title,Sifra,Vid,Litri,Temp,Zabeleska,Den,Reon,ReonIme,Koj,Vreme,LokalenId,Status)";
      if(den) q += "&$filter=fields/Den eq '" + den + "'";
      return site(l + q, { headers: { "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly" } });
    }).then(function(redovi){ return redovi.map(function(r){ var f = r.fields || {}; f.id = r.id; return f; }); });
  }
  function postaviStatus(itemId, status){
    return lista(C.listaVnesovi).then(function(l){ return graph(l + "/items/" + itemId + "/fields", { method: "PATCH", body: { Status: status } }); });
  }

  // Создавање на листите (само setup.html)
  function sozdadiListi(){
    return siteId().then(function(id){
      var b = "/sites/" + id + "/lists";
      var kolK = [
        { name: "Sifra", text: {}, indexed: true, enforceUniqueValues: true },
        { name: "Vid", choice: { choices: ["kravjo", "ovco", "kozjo"], displayAs: "dropDownMenu" } },
        { name: "Reon", text: {} }, { name: "ReonBr", number: {} }, { name: "Mesto", text: {} }, { name: "Otk", text: {} },
        { name: "Aktiven", boolean: {}, defaultValue: { value: "1" } }, { name: "Prethodna", text: {} }
      ];
      var kolV = [
        { name: "Sifra", text: {}, indexed: true }, { name: "Vid", text: {} }, { name: "Litri", number: { decimalPlaces: "one" } },
        { name: "Temp", text: {} }, { name: "Zabeleska", text: {} }, { name: "Den", text: {}, indexed: true }, { name: "Reon", text: {} },
        { name: "ReonIme", text: {} }, { name: "Koj", text: {} }, { name: "Vreme", text: {} },
        { name: "LokalenId", text: {}, indexed: true, enforceUniqueValues: true },
        { name: "Status", choice: { choices: ["Примено", "Потврдено", "Одбиено"], displayAs: "dropDownMenu" }, defaultValue: { value: "Примено" } }
      ];
      function napravi(ime, opis, koloni){
        return graph(b + "/" + encodeURIComponent(ime)).then(function(){ return "постои"; }).catch(function(e){
          if(e.status !== 404) throw e;
          return graph(b, { method: "POST", body: { displayName: ime, description: opis, columns: koloni, list: { template: "genericList" } } }).then(function(){ return "создадена"; });
        });
      }
      return napravi(C.listaKooperanti, "Кооперанти и лабораториски шифри (Откуп на терен)", kolK).then(function(a){
        return napravi(C.listaVnesovi, "Внесови на откупено млеко од терен", kolV).then(function(b2){ return { kooperanti: a, vnesovi: b2 }; });
      });
    });
  }

  window.Otkup = { init: init, smetka: smetka, najavi: najavi, odjavi: odjavi, graph: graph, kooperanti: kooperanti, dodajKooperant: dodajKooperant,
    pratiVnes: pratiVnes, vnesovi: vnesovi, postaviStatus: postaviStatus, sozdadiListi: sozdadiListi, config: C };
})();
