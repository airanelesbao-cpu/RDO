/* Ponte com o app nativo (Android/iPhone empacotado com Capacitor).
   No navegador comum este arquivo não faz nada. */
(function () {
  var C = window.Capacitor;
  if (!C) return;
  var plat = typeof C.getPlatform === 'function' ? C.getPlatform() : (C.platform || 'web');
  var nativo = typeof C.isNativePlatform === 'function' ? C.isNativePlatform() : (plat === 'android' || plat === 'ios');
  if (!nativo) return;
  window.RDO_NATIVE = true;
  var P = function (n) {
    if (typeof C.registerPlugin === 'function') return C.registerPlugin(n);
    return (C.Plugins || {})[n];
  };

  // 1) Voltou o sinal: sincroniza na hora (fotos e textos pendentes).
  try {
    P('Network').addListener('networkStatusChange', function (s) {
      if (s && s.connected && window.__rdo && window.__rdo.tick) window.__rdo.tick();
    });
  } catch (e) {}

  // 2) App foi para o segundo plano: termina o envio pendente antes de o sistema suspender.
  try {
    var App = P('App'), BG = P('BackgroundTask');
    App.addListener('appStateChange', async function (st) {
      if (st && st.isActive) { if (window.__rdo && window.__rdo.tick) window.__rdo.tick(); return; }
      var id = await BG.beforeExit(async function () {
        try { if (window.__rdo && window.__rdo.drain) await window.__rdo.drain(); } catch (e) {}
        BG.finish({ taskId: id });
      });
    });
  } catch (e) {}

  // 3) Foto da câmera salva na galeria do aparelho, sem perguntar.
  try {
    var Media = P('Media'), album = null;
    var achar = async function () {
      var r = await Media.getAlbums();
      var f = (r.albums || []).find(function (a) { return a.name === 'RDO'; });
      return f ? f.identifier : null;
    };
    var garantirAlbum = async function () {
      if (album) return album;
      try { album = await achar(); } catch (e) {}
      if (!album) { try { await Media.createAlbum({ name: 'RDO' }); album = await achar(); } catch (e) {} }
      return album;
    };
    window.rdoNativeSave = async function (dataUrl, nome) {
      var opts = { path: dataUrl, fileName: nome };
      if (plat === 'android') { var id = await garantirAlbum(); if (id) opts.albumIdentifier = id; }
      await Media.savePhoto(opts);
    };
  } catch (e) {}

  // 4) PDF/Excel: abre a folha de compartilhar do sistema (Arquivos, Drive, WhatsApp, e-mail...).
  try {
    var FS = P('Filesystem'), Share = P('Share');
    var paraB64 = function (blob) {
      return new Promise(function (ok, erro) {
        var r = new FileReader();
        r.onload = function () { ok(String(r.result).split(',')[1]); };
        r.onerror = erro;
        r.readAsDataURL(blob);
      });
    };
    window.rdoNativeShare = async function (blob, nome) {
      var w = await FS.writeFile({ path: nome, data: await paraB64(blob), directory: 'CACHE' });
      await Share.share({ title: nome, url: w.uri, dialogTitle: 'Salvar ou enviar' });
    };
  } catch (e) {}
})();
