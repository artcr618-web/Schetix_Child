#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Собирает Календарь/Календарь.html из Календарь/Календарь_дел.md.
Июль-стиль: большой календарь на месяц → клик по дню → список дел дня.
Звёздочка ★ = главное событие дня (показывается в ячейке месяца). Галочка ✓ = выполнено.
Отметки хранятся в браузере (localStorage).
Запуск: python3 Календарь/build_calendar.py
"""
import re, html, urllib.parse, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MD = os.path.join(ROOT, 'Календарь', 'Календарь_дел.md')
OUT = os.path.join(ROOT, 'Календарь', 'Календарь.html')

def esc(t):
    return html.escape(t, quote=False)

def md_marks(t):
    t = esc(t)
    t = re.sub(r'==([^=]+)==', r'<mark>\1</mark>', t)
    return t

def load_doc(path):
    try:
        with open(os.path.join(ROOT, path), encoding='utf-8') as f:
            return f.read()
    except Exception:
        return '— текст не найден: %s —' % path

# ---------- разбор плана ----------
with open(MD, encoding='utf-8') as f:
    lines = f.read().split('\n')

days, top_notes = [], []
cur_day = cur_sub = cur_item = None

for ln in lines:
    if ln.startswith('> '):
        (top_notes if cur_day is None else cur_day['notes']).append(ln[2:].strip())
        continue
    if ln.startswith('## '):
        title = ln[3:].strip()
        m = re.match(r'([А-Яа-я]+),\s*(\d{2}\.\d{2})', title)
        cur_day = {'title': title, 'weekday': m.group(1) if m else '',
                   'date': m.group(2) if m else '', 'notes': [], 'subs': []}
        days.append(cur_day); cur_sub = cur_item = None
        continue
    if ln.startswith('### '):
        cur_sub = {'title': ln[4:].strip(), 'items': []}
        if cur_day is None:
            cur_day = {'title': '', 'weekday': '', 'date': '', 'notes': [], 'subs': []}
            days.append(cur_day)
        cur_day['subs'].append(cur_sub); cur_item = None
        continue
    m = re.match(r'(?:\d+\.|-)\s+\[([ x])\]\s+(.*)$', ln)
    if m:
        rest = m.group(2).strip()
        mb = re.match(r'\*\*(.+?)\*\*\s*(.*)$', rest)
        if mb:
            title, sub = mb.group(1).strip(), mb.group(2).strip()
        else:
            title, sub = rest, ''
        item = {'title': title, 'sub': sub,
                'bullets': [], 'file': None, 'question': False}
        if cur_sub is None:
            cur_sub = {'title': '', 'items': []}
            if cur_day is None:
                cur_day = {'title': '', 'weekday': '', 'date': '', 'notes': [], 'subs': []}
                days.append(cur_day)
            cur_day['subs'].append(cur_sub)
        if cur_sub['title'].upper().startswith('ПОД ВОПРОСОМ'):
            item['question'] = True
        cur_sub['items'].append(item); cur_item = item
        continue
    if ln.startswith('    - ') and cur_item is not None:
        b = ln[6:].strip()
        if b.upper().startswith('ФАЙЛ:'):
            cur_item['file'] = b.split(':', 1)[1].strip()
        else:
            cur_item['bullets'].append(b)
        continue

# исторический день: 15.10 — увоз
days.append({'title': '15.10 — УВОЗ ДОЧЕРИ ИЗ ШКОЛЫ (событие)', 'weekday': 'Чт', 'date': '15.10',
             'notes': [], 'subs': [{'title': '', 'items': [{
                 'title': '⚠ Увоз дочери из школы отцом (событие дня)', 'sub': '',
                 'bullets': ['ФАКТ: забрана с последнего урока (урок алгебры) под ложным предлогом «по запросу опеки»; хронология звонков — в Доказательствах'],
                 'file': 'История/События/2026-10-15_Увоз_дочери_из_школы.md', 'question': False}]}]})
days.sort(key=lambda d: (int(d['date'][-2:]) if d['date'] else 99))

# ---------- месяц ----------
WD_SHORT = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
MONTH_DAYS = 31
FIRST_WD = 3  # 01.10.2026 — четверг

def datekey(dstr):
    return '10' + dstr[:2] if dstr else ''

by_key = {datekey(d['date']): d for d in days if d['date']}

def badge_for(sub_title):
    m = re.search(r'(\d{1,2}:\d{2}\s*[–-]\s*\d{1,2}:\d{2}|\d{1,2}:\d{2})', sub_title)
    return m.group(1) if m else ''

def item_meta(item, sub_title):
    badge = badge_for(sub_title)
    for b in item['bullets']:
        if b.upper().startswith('КОГДА'):
            badge = b.split(':', 1)[1].strip()[:16] or badge
            break
    where = ''
    for b in item['bullets']:
        if b.upper().startswith(('ГДЕ', 'КУДА')):
            where = b.split(':', 1)[1].strip()
            break
    return badge, where

def is_ev(it):
    return ('СОБЫТИЕ' in it['title'].upper()) or it['title'].startswith(('⚖', '⚠', '⏰'))

def bullet_html(b):
    if ':' in b:
        k, v = b.split(':', 1)
        if len(k) <= 22 and not k.startswith('http'):
            return '<li><b>%s:</b>%s</li>' % (esc(k.strip()), md_marks(v))
    return '<li>%s</li>' % md_marks(b)

css = '''
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Google Sans', Roboto, Arial, sans-serif; background: #f8f9fa; color: #3c4043; }
.topbar { background: #fff; border-bottom: 1px solid #e0e0e0; padding: 10px 18px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
.logo { font-size: 20px; color: #5f6368; font-weight: 500; }
.logo b { color: #1a73e8; }
.btn { background: #1a73e8; color: #fff; border: none; border-radius: 24px; padding: 9px 18px; font-size: 13px; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; }
.btn:hover { background: #1765cc; }
.btn.grey { background: #fff; color: #1a73e8; border: 1px solid #dadce0; }
.btn.grey:hover { background: #f1f3f4; }
.alert { background: #fef7e0; border: 1px solid #f9d77e; border-radius: 10px; padding: 12px 16px; margin: 14px 18px 0; font-size: 13px; color: #7c6221; }
.wrap { max-width: 1150px; margin: 0 auto; padding: 16px 18px 48px; }
.month-title { font-size: 22px; font-weight: 500; margin: 12px 2px 14px; display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
.month-title .hint { font-size: 12px; color: #80868b; font-weight: 400; }
.grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 6px; }
.wd { text-align: center; font-size: 11px; letter-spacing: 1px; text-transform: uppercase; color: #80868b; padding: 4px 0; }
.cell { background: #fff; border: 1px solid #e0e0e0; border-radius: 12px; min-height: 104px; padding: 7px 8px; cursor: pointer; transition: box-shadow .15s; position: relative; }
.cell:hover { box-shadow: 0 1px 6px rgba(60,64,67,.25); }
.cell.empty { background: #f1f3f4; border-color: #f1f3f4; cursor: default; }
.cell.sel { border: 2px solid #1a73e8; }
.cell .num { width: 26px; height: 26px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 13px; font-weight: 500; }
.cell.today .num { background: #1a73e8; color: #fff; }
.cell.hist .num { background: #d93025; color: #fff; }
.cell.ev .num { background: #f9ab00; color: #fff; }
.item.event { border-left: 5px solid #f9ab00; background: #fffdf5; }
.cell .cnt { font-size: 10.5px; color: #1a73e8; background: #e8f0fe; border-radius: 10px; padding: 1px 7px; margin-top: 3px; display: inline-block; }
.cell .evline { margin-top: 5px; font-size: 10.5px; color: #a14a00; line-height: 1.35; font-weight: 500; }
.cell .starline { margin-top: 5px; font-size: 10.5px; color: #a14a00; line-height: 1.35; }
.cell .starline b { color: #f9ab00; }
.daypanel { display: none; background: #fff; border: 1px solid #e0e0e0; border-radius: 16px; margin-top: 18px; padding: 18px 20px; }
.daypanel.open { display: block; }
.day-title { font-size: 18px; font-weight: 500; margin-bottom: 4px; display: flex; gap: 10px; align-items: baseline; flex-wrap: wrap; }
.day-title .dt { color: #1a73e8; }
.day-title .sm { font-size: 13px; color: #5f6368; font-weight: 400; }
.day-note { background: #e8f0fe; border-radius: 10px; padding: 10px 14px; font-size: 12.5px; color: #174ea6; margin: 8px 0; line-height: 1.5; }
.sub-title { font-size: 11.5px; font-weight: 600; letter-spacing: .8px; text-transform: uppercase; color: #5f6368; margin: 16px 0 8px; }
.item { background: #fff; border: 1px solid #e0e0e0; border-radius: 12px; margin-bottom: 8px; overflow: hidden; }
.item.question { border-left: 5px solid #f9ab00; background: #fffdf5; }
.item-hd { padding: 11px 14px; display: flex; align-items: center; gap: 10px; cursor: pointer; }
.item-hd:hover { background: #f8f9fa; }
.item.question .item-hd:hover { background: #fef9e7; }
.badge { background: #1a73e8; color: #fff; font-size: 11px; padding: 3px 8px; border-radius: 5px; white-space: nowrap; }
.badge.grey { background: #f1f3f4; color: #5f6368; }
.badge.amber { background: #f9ab00; color: #fff; }
.item-title { flex: 1; font-size: 14px; font-weight: 500; }
.item-where { font-size: 11.5px; color: #5f6368; }
.chev { transition: transform .15s; color: #5f6368; font-size: 11px; }
.item.open .chev { transform: rotate(180deg); }
.item-body { display: none; padding: 0 14px 14px 14px; border-top: 1px solid #f1f3f4; }
.item.open .item-body { display: block; }
.item.done-hidden { display: none; }
.showdone { margin: 8px 0 0; }
.bullets { margin: 12px 0 8px; }
.bullets li { list-style: none; padding: 4px 0 4px 18px; position: relative; font-size: 12.5px; line-height: 1.55; }
.bullets li:before { content: '•'; position: absolute; left: 2px; color: #1a73e8; }
.bullets li b { color: #202124; }
.prev-wrap { margin-top: 10px; border: 1px solid #e0e0e0; border-radius: 10px; overflow: hidden; }
.prev-hd { background: #f8f9fa; padding: 8px 12px; font-size: 12px; color: #5f6368; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.prev-hd .name { font-weight: 500; color: #3c4043; flex: 1; min-width: 140px; word-break: break-all; }
.prev { padding: 12px 14px; font-size: 12px; line-height: 1.6; white-space: pre-wrap; max-height: 320px; overflow-y: auto; }
mark { background: #fff3c2; padding: 0 2px; border-radius: 2px; }
.hint { font-size: 11.5px; color: #80868b; margin: 4px 2px; }
.chip { display: inline-flex; align-items: center; gap: 5px; border: 1px solid #dadce0; background: #fff; color: #5f6368; border-radius: 20px; padding: 4px 11px; font-size: 11.5px; cursor: pointer; user-select: none; }
.chip:hover { background: #f1f3f4; }
.chip.on { background: #e6f4ea; border-color: #34a853; color: #137333; }
.chip.star.on { background: #fef7e0; border-color: #f9ab00; color: #a14a00; }
.chip.star.on .s { color: #f9ab00; }
.legend { display: flex; gap: 12px; flex-wrap: wrap; margin: 10px 2px 0; }
'''

# ---------- HTML ----------
parts = ['''<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Календарь — октябрь 2026</title><style>%s</style></head>
<body>
<div class="topbar">
  <div class="logo"><b>Календарь</b> — октябрь 2026</div>
  <a class="btn grey" href="%s">✏ Редактировать план дня</a>
  <span class="hint">клик по дню — список дел; ✓ — готово; ★ — главное событие дня</span>
</div>
<div class="wrap">
<div class="month-title">Октябрь 2026
  <span class="hint">клик по дню → дела дня; звёздочка ★ в деле → главное событие (видно в календаре)</span>
</div>
<div class="grid">''' % (css, urllib.parse.quote('../Календарь/Календарь_дел.md'))]

for w in WD_SHORT:
    parts.append('<div class="wd">%s</div>' % w)
cells_html = ''
for i in range(MONTH_DAYS):
    dnum = i + 1
    pos = FIRST_WD + i
    if i == 0:
        for _ in range(FIRST_WD):
            cells_html += '<div class="cell empty"></div>'
    key = '10%02d' % dnum
    day = by_key.get(key)
    cls = 'cell'
    if dnum == 17:
        cls += ' today'
    if dnum == 15 and day:
        cls += ' hist'
    if dnum == 20 and day:
        cls += ' ev'
    inner = '<div class="num">%d</div>' % dnum
    if day:
        evs = [it for s in day['subs'] for it in s['items'] if is_ev(it)]
        for ev in evs[:3]:
            short = ev['title'].replace('(СОБЫТИЕ)', '').replace('СОБЫТИЕ', '').strip()
            short = re.sub(r'^[⚖⚠⏰]\s*', '', short)
            if len(short) > 26:
                short = short[:25] + '…'
            inner += '<div class="evline">★ %s</div>' % esc(short)
        n_zad = sum(1 for s in day['subs'] for it in s['items'] if not is_ev(it))
        if n_zad:
            inner += '<div class="cnt">%d задач</div>' % n_zad
        inner += '<div class="starline" id="star-%s"></div>' % key
    cells_html += '<div class="%s" id="cell-%s" onclick="openDay(\'%s\')">%s</div>' % (cls, key, key, inner)
parts.append(cells_html)
parts.append('</div>')
parts.append('<div class="legend"><span class="chip">✓ готово (домашка)</span><span class="chip star on"><span class="s">★</span> главное событие дня</span><span class="chip">отметки сохраняются в браузере</span></div>')

# панели дней
def render_panel(key, day):
    pid = 'panel-' + key
    title = day['title']
    parts.append('<div class="daypanel" id="%s">' % pid)
    parts.append('<button class="chip showdone" style="display:none" onclick="toggleShowDone(\'%s\')"></button>' % pid)
    parts.append('<div class="day-title"><span class="dt">%s</span><span>%s</span></div>' % (
        esc((day['weekday'] + ', ' + day['date']).strip(', ')),
        esc(title.split('—', 1)[1].strip() if '—' in title else '')))
    for n in day['notes']:
        parts.append('<div class="day-note">%s</div>' % md_marks(n))
    idx = 0
    for sub in day['subs']:
        if sub['title']:
            parts.append('<div class="sub-title">%s</div>' % esc(sub['title']))
        for it in sub['items']:
            idx += 1
            tid = '%s-t%d' % (key, idx)
            badge, where = item_meta(it, sub['title'])
            is_event = is_ev(it)
            is_q = it['question'] or 'ПОД ВОПРОСОМ' in it['title'].upper()
            cls = 'item question' if is_q else ('item event' if is_event else 'item')
            if is_event:
                badge, bcls = 'СОБЫТИЕ', 'badge amber'
            else:
                badge = 'задача'
                bcls = 'badge grey'
            qtag = '<span class="badge amber">под вопросом</span>' if is_q else ''
            parts.append('<div class="%s" id="item-%s">' % (cls, tid))
            parts.append('<div class="item-hd" onclick="if(!event.target.closest(\'.chip\'))this.parentNode.classList.toggle(\'open\')">')
            parts.append('<span class="chip done" id="done-%s" onclick="event.stopPropagation();toggleDone(\'%s\',this)" title="домашка готова">✓ готово</span>' % (tid, tid))
            parts.append('<span class="chip star" id="star-%s" onclick="event.stopPropagation();toggleStar(\'%s\',\'%s\',this)" title="главное событие дня"><span class="s">★</span></span>' % (tid, key, tid))
            parts.append('<span class="%s">%s</span>' % (bcls, esc(badge or '—')))
            parts.append('<div class="item-title">%s %s</div>' % (esc(it['title']), qtag))
            if where:
                parts.append('<div class="item-where">%s</div>' % esc(where.split(',')[0][:24]))
            parts.append('<span class="chev">▼</span></div><div class="item-body">')
            if it['sub']:
                parts.append('<div class="hint">%s</div>' % md_marks(it['sub']))
            if it['bullets']:
                parts.append('<ul class="bullets">' + ''.join(bullet_html(b) for b in it['bullets']) + '</ul>')
            if it['file']:
                rel = it['file']
                href = urllib.parse.quote('../' + rel)
                text = load_doc(rel)
                body = '\n'.join(l for l in text.split('\n') if not l.startswith('> ')).strip()
                parts.append('''<div class="prev-wrap">
<div class="prev-hd"><span class="name">%s</span>
<a class="btn grey" style="padding:5px 12px;font-size:12px" href="%s">✏ Редактировать</a></div>
<div class="prev">%s</div>
<div class="hint" style="padding:6px 12px">файл: %s — правки в .md; «обнови календарь» пересоберёт превью</div>
</div>''' % (esc(os.path.basename(rel)), href, md_marks(body), esc(rel)))
            parts.append('</div></div>')
    parts.append('</div>')

for key, day in sorted(by_key.items()):
    render_panel(key, day)

# пустая панель (дни без плана)
parts.append('<div class="daypanel" id="panel-empty"><div class="day-title"><span class="dt" id="empty-date"></span><span>Событий нет</span></div></div>')
parts.append('</div>')  # /wrap

parts.append('''<script>
function $(id){ return document.getElementById(id); }
function lsGet(k, d){ try { var v = JSON.parse(localStorage.getItem(k)); return v === null ? d : v; } catch(e){ return d; } }
function lsSet(k, v){ try { localStorage.setItem(k, JSON.stringify(v)); } catch(e){} }

function openDay(key){
  document.querySelectorAll('.daypanel').forEach(function(p){ p.classList.remove('open'); });
  document.querySelectorAll('.cell').forEach(function(c){ c.classList.remove('sel'); });
  var panel = $('panel-' + key);
  if (panel) { panel.classList.add('open'); }
  else {
    $('panel-empty').classList.add('open');
    var n = key.slice(2), t = $('empty-date'); if (t) t.textContent = n + '.10.2026';
  }
  var cell = $('cell-' + key); if (cell) cell.classList.add('sel');
  lsSet('cal_last_day', key);
  if (panel) panel.scrollIntoView({behavior:'smooth', block:'start'});
}

function toggleDone(tid, el){
  el.classList.toggle('on');
  var on = el.classList.contains('on');
  lsSet('done_' + tid, on);
  var item = $('item-' + tid);
  if (item) {
    var panel = item.closest('.daypanel');
    var show = panel && panel.classList.contains('show-completed');
    if (on && !show) item.classList.add('done-hidden');
    else item.classList.remove('done-hidden');
  }
  refreshShowBtns();
}
function refreshShowBtns(){
  document.querySelectorAll('.daypanel').forEach(function(panel){
    var hidden = panel.querySelectorAll('.item.done-hidden').length;
    var btn = panel.querySelector('.showdone');
    if (!btn) return;
    if (hidden > 0) { btn.style.display = 'inline-flex'; btn.textContent = '☑ показать выполненные (' + hidden + ')'; }
    else btn.style.display = 'none';
  });
}
function toggleShowDone(panelId){
  var panel = $(panelId);
  panel.classList.toggle('show-completed');
  if (panel.classList.contains('show-completed')) {
    panel.querySelectorAll('.item.done-hidden').forEach(function(i){ i.classList.remove('done-hidden'); });
  } else {
    panel.querySelectorAll('.item').forEach(function(i){
      var chip = i.querySelector('.chip.done');
      if (chip && chip.classList.contains('on')) i.classList.add('done-hidden');
    });
  }
  refreshShowBtns();
}

function toggleStar(dayKey, tid, el){
  var titles = lsGet('cal_stars', {});
  var wasOn = el.classList.contains('on');
  document.querySelectorAll('.item .chip.star').forEach(function(s){
    if (s.closest('.daypanel') === $('panel-' + dayKey) || s.closest('.daypanel') === null) {
      if (s.id.indexOf('star-' + dayKey + '-') === 0) s.classList.remove('on');
    }
  });
  if (!wasOn) {
    el.classList.add('on');
    var item = $('item-' + tid);
    var t = item ? item.querySelector('.item-title').textContent.trim() : '';
    titles[dayKey] = { id: tid, title: t.slice(0, 46) };
  } else {
    delete titles[dayKey];
  }
  lsSet('cal_stars', titles);
  refreshCell(dayKey);
}

function refreshCell(dayKey){
  var titles = lsGet('cal_stars', {});
  var box = $('star-' + dayKey);
  if (!box) return;
  if (titles[dayKey]) box.innerHTML = '<b>★</b> ' + titles[dayKey].title;
  else box.innerHTML = '';
}

function init(){
  document.querySelectorAll('[id^="done-"]').forEach(function(el){
    var tid = el.id.slice(5);
    if (lsGet('done_' + tid, false)) {
      el.classList.add('on');
      var it = $('item-' + tid);
      if (it) it.classList.add('done-hidden');
    }
  });
  refreshShowBtns();
  var titles = lsGet('cal_stars', {});
  Object.keys(titles).forEach(function(dayKey){
    var st = $('star-' + titles[dayKey].id);
    if (st) st.classList.add('on');
    refreshCell(dayKey);
  });
  var last = lsGet('cal_last_day', '1017');
  openDay($('panel-' + last) || $('cell-' + last) ? last : '1017');
}
init();
</script>
</body></html>''')

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(parts))
print('OK ->', OUT)
