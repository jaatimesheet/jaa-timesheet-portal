// JAA Timesheet — main.js

function toast(msg, type='success') {
  const icons = {success:'✓', error:'✕', warning:'⚠', info:'ℹ'};
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.innerHTML = `<span>${icons[type]||'•'}</span><span>${msg}</span>`;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => {
    el.style.cssText = 'opacity:0;transform:translateX(20px);transition:all .3s';
    setTimeout(() => el.remove(), 350);
  }, 3200);
}

function openModal(id)  { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

document.addEventListener('click', e => {
  if (e.target.classList.contains('modal-overlay'))
    e.target.classList.remove('open');
});

// Hours badge helper
function hrsBadge(hrs, target) {
  const pct = hrs / target;
  let cls, icon, label;
  if (hrs >= target)  { cls='complete'; icon='🟢'; label='Complete'; }
  else if (pct >= .5) { cls='partial';  icon='🟡'; label='Partial'; }
  else                { cls='low';      icon='🔴'; label='Incomplete'; }
  return `<span class="hrs-badge ${cls}">${icon} ${hrs.toFixed(1)} / ${target} h — ${label}</span>`;
}

function progressBar(hrs, target) {
  const pct = Math.min(hrs / target, 1);
  const cls = hrs >= target ? 'done' : (pct >= .5 ? 'warn' : 'low');
  return `<div class="progress-wrap"><div class="progress-fill ${cls}" style="width:${(pct*100).toFixed(1)}%"></div></div>`;
}

// Async API helper
async function api(url, data) {
  const r = await fetch(url, {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify(data)
  });
  return r.json();
}

// ── Session timeout warning ───────────────────────────────────────────────
// Reads SESSION_TIMEOUT_MINS injected by timesheet page; warns 5 min before expiry
(function() {
  const el = document.getElementById('session-timeout-mins');
  if (!el) return;
  const mins = parseInt(el.dataset.mins || '60');
  const warnAt = (mins - 5) * 60 * 1000;
  const logoutAt = mins * 60 * 1000;
  let warnTimer  = setTimeout(() => {
    toast('⚠ Session expires in 5 minutes due to inactivity.', 'warning');
  }, warnAt);
  let logoutTimer = setTimeout(() => {
    window.location.href = '/logout';
  }, logoutAt);
  // Reset on activity
  ['click','keydown','scroll','mousemove'].forEach(ev => {
    document.addEventListener(ev, () => {
      clearTimeout(warnTimer); clearTimeout(logoutTimer);
      warnTimer   = setTimeout(() => toast('⚠ Session expires in 5 minutes due to inactivity.','warning'), warnAt);
      logoutTimer = setTimeout(() => { window.location.href = '/logout'; }, logoutAt);
    }, {passive: true});
  });
})();
