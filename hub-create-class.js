/**
 * Skill Hub — "Create Class" for teachers.
 *
 * One form creates the class in every app the teacher ticks, then saves the class codes in the
 * Master Sheet (through the hub Apps Script, action saveClass) under one HUB code that students
 * type in "Create Account".
 *
 * The teacher signs in with the SAME email + password in every app (each app has its own
 * Firebase project, so each sign-in uses its own named Firebase app instance).
 *
 * The class documents below mirror what each app writes itself when a teacher presses
 * "Create class" in that app (see createClass / classCreate in the app's fbdata.js). If an
 * app's class schema changes, update the matching `doc` builder here.
 *
 * Class Hub has no Firestore "create class": its classes come from the Class Hub sheet. So for
 * Class Hub the teacher picks a class that already exists there.
 *
 * Needs on the page: firebase-app/auth/firestore/database compat SDKs, window.HUB_API, and the
 * Class Hub data layer (window.DB, from class-hub/db.js) for the class list.
 */
(function () {
  'use strict';

  var FB_CONFIG = {
    speaking: {
      apiKey: 'AIzaSyAeB-tcXD9QOkppW4oshpFI4aXe9T33kws', authDomain: 'fluentalk-f2ed4.firebaseapp.com',
      projectId: 'fluentalk-f2ed4', storageBucket: 'fluentalk-f2ed4.firebasestorage.app',
      messagingSenderId: '298459732541', appId: '1:298459732541:web:c9372297de253f7f01e854'
    },
    listening: {
      apiKey: 'AIzaSyABj5BoT_Bz8aGJ6bys8LWCLAFhut5VJL8', authDomain: 'listendictation-4c26e.firebaseapp.com',
      projectId: 'listendictation-4c26e', storageBucket: 'listendictation-4c26e.firebasestorage.app',
      messagingSenderId: '76602878721', appId: '1:76602878721:web:67c0b5c192c6ed88828900'
    },
    vocab: {
      apiKey: 'AIzaSyAhoWEygnchnXPf1BaG2T6ZvpV0VY7oeeY', authDomain: 'vocabmaster-3a0dd.firebaseapp.com',
      projectId: 'vocabmaster-3a0dd', storageBucket: 'vocabmaster-3a0dd.firebasestorage.app',
      messagingSenderId: '199135021961', appId: '1:199135021961:web:7e6bd1bca219b1ab32af58'
    },
    writing: {
      apiKey: 'AIzaSyCj8WTr6eaqMGhqKltiZ9444LELV-7ZDIw', authDomain: 'articuwrite.firebaseapp.com',
      databaseURL: 'https://articuwrite-default-rtdb.asia-southeast1.firebasedatabase.app',
      projectId: 'articuwrite'
    }
  };

  var LABEL = { speaking: 'Speaking', listening: 'Listening', vocab: 'Vocab', writing: 'Writing', classhub: 'Class Hub' };
  var ORDER = ['speaking', 'listening', 'vocab', 'writing'];

  var L3 = 'ABCDEFGHJKMNPQRSTUVWXYZ', D3 = '23456789', C5 = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  function pick(set, n) {
    var a = new Uint8Array(n), out = '';
    crypto.getRandomValues(a);
    for (var i = 0; i < n; i++) out += set[a[i] % set.length];
    return out;
  }
  // Same code shapes as each app's own Create class
  var GEN = {
    speaking : function () { return pick(L3, 3) + pick(D3, 3); },
    listening: function () { return 'LD-' + pick(L3, 3) + pick(D3, 3); },
    vocab    : function () { return 'VM-' + pick(C5, 5); },
    writing  : function () { return 'AW-' + pick(C5, 5); }
  };

  // Document each app stores in classes/{code}
  var DOC = {
    speaking: function (c, f) {
      return { ClassID: c, ClassName: f.className, AcademicYear: f.year, Semester: f.semester,
               TeacherName: f.teacherName, Status: 'Active', CreatedAt: f.now };
    },
    listening: function (c, f) {
      return { classId: c, className: f.className, academicYear: f.year, semester: f.semester,
               teacherUid: f.uid, teacherName: f.teacherName, teacherEmail: f.email, status: 'Active', createdAt: f.now };
    },
    vocab: function (c, f) {
      return { classId: c, className: f.className, academicYear: f.year, semester: f.semester,
               teacherUid: f.uid, teacherName: f.teacherName, teacherEmail: f.email, status: 'Active', createdAt: f.now };
    },
    writing: function (c, f) {
      return { classId: c, className: f.className, year: f.year, semester: f.semester,
               teacherUid: f.uid, teacherEmail: f.email, aiEnabled: true, archived: false, createdAt: f.now };
    }
  };

  // Firebase refuses passwords under 6 chars; every app pads them the same way
  function authPw(p) { p = String(p == null ? '' : p); return p.length >= 6 ? p : (p + '______').slice(0, 6); }

  var apps = {};
  function fbApp(key) {
    if (apps[key]) return apps[key];
    var app = firebase.initializeApp(FB_CONFIG[key], 'hub-' + key);
    var auth = app.auth();
    try { auth.setPersistence(firebase.auth.Auth.Persistence.NONE); } catch (e) {}
    apps[key] = { app: app, auth: auth, fs: app.firestore() };
    return apps[key];
  }

  function nice(e) {
    var c = (e && e.code) || '';
    if (/invalid-credential|wrong-password|user-not-found|invalid-login/.test(c)) return 'wrong email or password for this app';
    if (/too-many-requests/.test(c)) return 'too many attempts, wait a few minutes';
    if (/permission-denied/.test(c)) return 'permission denied (is this a teacher account of this app?)';
    return (e && e.message) || String(e);
  }

  // Sign in as teacher in one app, create the class there. → { code, idToken }
  async function createInApp(key, f) {
    var h = fbApp(key);
    // Speaking may use its own teacher account
    var em = (key === 'speaking' && f.spEmail) ? f.spEmail : f.email;
    var pw = (key === 'speaking' && f.spEmail) ? f.spPassword : f.password;
    var cred = await h.auth.signInWithEmailAndPassword(em, authPw(pw));
    var uid = cred.user.uid;
    var idToken = await cred.user.getIdToken();

    var teacherName = '';
    try {
      var u = await h.fs.doc('users/' + uid).get();
      if (u.exists) teacherName = u.data().fullName || u.data().name || '';
    } catch (e) {}

    var form = { className: f.className, year: f.year, semester: f.semester, uid: uid,
                 email: em, teacherName: teacherName, now: new Date().toISOString() };
    var code = '';
    for (var i = 0; i < 20; i++) {
      var c = GEN[key]();
      if (!(await h.fs.doc('classes/' + c).get()).exists) { code = c; break; }
    }
    if (!code) throw new Error('could not find a free class code');
    await h.fs.doc('classes/' + code).set(DOC[key](code, form));

    if (key === 'writing') {
      // lets this teacher watch the class Live (same as ArticuWrite's own Create class)
      try { await h.app.database().ref('classOwner/' + code).set(uid); }
      catch (e) { return { code: code, idToken: idToken, warn: 'class created, but Live view could not be enabled: ' + nice(e) }; }
    }
    return { code: code, idToken: idToken };
  }

  function post(body) {
    return fetch(window.HUB_API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) })
      .then(function (r) { return r.json(); });
  }

  async function run(f) {
    var results = {}, codes = {}, tokenFrom = null, token = '';
    for (var i = 0; i < f.apps.length; i++) {
      var key = f.apps[i];
      try {
        var r = await createInApp(key, f);
        results[key] = { ok: true, code: r.code, warn: r.warn || '' };
        codes[key] = r.code;
        if (!token) { token = r.idToken; tokenFrom = key; }
      } catch (e) {
        results[key] = { ok: false, error: nice(e) };
      }
    }
    if (f.classHubName) { codes.classhub = f.classHubName; results.classhub = { ok: true, code: f.classHubName }; }

    var hubCode = '';
    if (token) {
      try {
        var res = await post({ action: 'saveClass', idToken: token, project: tokenFrom, className: f.className,
                               codes: codes, classHubName: f.classHubName || '', hubCode: f.existing || '' });
        if (res.success) hubCode = res.hubCode; else results._save = { ok: false, error: res.error };
      } catch (e) { results._save = { ok: false, error: 'could not reach the hub server' }; }
    } else {
      results._save = { ok: false, error: 'no app created a class, nothing saved' };
    }
    Object.keys(apps).forEach(function (k) { try { apps[k].auth.signOut(); } catch (e) {} });
    return { results: results, hubCode: hubCode };
  }

  // ── UI ─────────────────────────────────────────────────────────────────────
  function $(id) { return document.getElementById(id); }
  function init() {
    var overlay = $('ccOverlay'), openBtn = $('openCreateClass');
    if (!overlay || !openBtn) return;
    var toast = $('ccToast'), submit = $('ccSubmit'), result = $('ccResult'), form = $('ccForm');

    function say(msg, type) {
      toast.textContent = msg; toast.className = 'su-toast' + (type ? ' ' + type : '');
      if (msg) { try { toast.scrollIntoView({ block: 'center', behavior: 'smooth' }); } catch (e) {} }
    }
    function close() { overlay.classList.remove('open'); }

    function loadClassHubClasses() {
      var sel = $('ccHubClass');
      if (sel.dataset.loaded || !window.DB) return;
      DB.publicClassNames().then(function (r) {
        sel.innerHTML = '<option value="">Select an existing Class Hub class…</option>' +
          (r.classes || []).map(function (n) { return '<option>' + String(n).replace(/</g, '&lt;') + '</option>'; }).join('');
        sel.dataset.loaded = '1';
      }).catch(function () {});
    }

    openBtn.onclick = function () {
      overlay.classList.add('open');
      form.style.display = ''; result.style.display = 'none'; say('');
      loadClassHubClasses();
    };
    $('ccClose').onclick = close;
    overlay.onclick = function (e) { if (e.target === overlay) close(); };

    var boxes = ORDER.concat(['classhub']).map(function (k) { return $('cc_' + k); });
    $('cc_all').onchange = function () {
      var on = this.checked;
      boxes.forEach(function (b) { b.checked = on; });
      $('ccHubWrap').style.display = on ? '' : 'none';
    };
    $('cc_classhub').onchange = function () { $('ccHubWrap').style.display = this.checked ? '' : 'none'; };
    $('cc_spDiff').onchange = function () { $('ccSpWrap').style.display = this.checked ? '' : 'none'; };

    submit.onclick = async function () {
      var chosen = ORDER.filter(function (k) { return $('cc_' + k).checked; });
      var wantHub = $('cc_classhub').checked;
      var f = {
        email: $('ccEmail').value.trim().toLowerCase(), password: $('ccPw').value,
        className: $('ccName').value.trim(), year: $('ccYear').value.trim(), semester: $('ccSem').value.trim(),
        apps: chosen, classHubName: wantHub ? $('ccHubClass').value : '',
        existing: $('ccExisting').value.trim().toUpperCase()
      };
      if ($('cc_spDiff').checked) { f.spEmail = $('ccSpEmail').value.trim().toLowerCase(); f.spPassword = $('ccSpPw').value; }
      if (!f.className) return say('⚠ Enter the class name.', 'err');
      if (!chosen.length) return say('⚠ Tick at least one of Speaking, Listening, Vocab or Writing (Class Hub alone does not need the hub).', 'err');
      if (wantHub && !f.classHubName) return say('⚠ Choose the Class Hub class, or untick Class Hub.', 'err');
      if (!f.email || !f.password) return say('⚠ Enter your teacher email and password.', 'err');
      if ($('cc_spDiff').checked && chosen.indexOf('speaking') >= 0 && (!f.spEmail || !f.spPassword))
        return say('⚠ Enter the Speaking teacher email and password.', 'err');
      if (f.existing && !/^HUB-[A-Z0-9]{6}$/.test(f.existing)) return say('⚠ The existing class code looks like HUB-ABC123.', 'err');

      submit.disabled = true; submit.innerHTML = '<span class="su-spin"></span> Creating class…'; say('');
      var out;
      try { out = await run(f); }
      catch (e) { out = { results: { _save: { ok: false, error: nice(e) } }, hubCode: '' }; }
      submit.disabled = false; submit.textContent = 'Create Class';

      form.style.display = 'none'; result.style.display = '';
      var lines = [];
      Object.keys(out.results).forEach(function (k) {
        var r = out.results[k];
        if (k === '_save') { if (!r.ok) lines.push('✗ Save to hub: ' + r.error); return; }
        lines.push(r.ok ? '✓ ' + LABEL[k] + ': ' + r.code + (r.warn ? '  (' + r.warn + ')' : '')
                        : '✗ ' + LABEL[k] + ': ' + r.error);
      });
      $('ccHubCode').textContent = out.hubCode || '—';
      $('ccHubCodeWrap').style.display = out.hubCode ? '' : 'none';
      $('ccLines').textContent = lines.join('\n');
    };
    $('ccAgain').onclick = function () { form.style.display = ''; result.style.display = 'none'; say(''); };
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
