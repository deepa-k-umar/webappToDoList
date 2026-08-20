const STORAGE_KEY = 'monthly-todo-tasks-v1';
const initialTasks = [{ id: '4e227f0a27444d968713738051547553', title: 'Airtel Postpaid bill payment monthly', month: '2026-08', frequency: 'Monthly', done: true }];
let tasks = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null') || initialTasks;
let selectedMonth = new Date().toISOString().slice(0, 7);
let activeFilter = 'all';

const $ = id => document.getElementById(id);
const monthLabel = value => new Intl.DateTimeFormat('en', { month: 'long', year: 'numeric' }).format(new Date(`${value}-01T12:00:00`));
const addMonths = (value, amount) => { const date = new Date(`${value}-01T12:00:00`); date.setMonth(date.getMonth() + amount); return date.toISOString().slice(0, 7); };
const save = () => localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks));

function render() {
  $('monthPicker').value = selectedMonth;
  $('monthName').textContent = monthLabel(selectedMonth);
  const query = $('searchInput').value.trim().toLowerCase();
  const monthTasks = tasks.filter(task => task.month === selectedMonth);
  const filtered = monthTasks.filter(task => (activeFilter === 'all' || (activeFilter === 'open' && !task.done) || (activeFilter === 'done' && task.done)) && task.title.toLowerCase().includes(query));
  const done = monthTasks.filter(task => task.done).length;
  $('totalCount').textContent = monthTasks.length;
  $('openCount').textContent = monthTasks.length - done;
  $('doneCount').textContent = done;
  $('listHeading').textContent = `${activeFilter === 'open' ? 'Open tasks' : activeFilter === 'done' ? 'Completed tasks' : 'Tasks'} · ${monthLabel(selectedMonth)}`;
  $('taskList').innerHTML = filtered.map((task, index) => `<article class="task ${task.done ? 'done' : ''}" style="animation-delay:${index * 35}ms"><input class="task-check" type="checkbox" ${task.done ? 'checked' : ''} data-id="${task.id}" aria-label="Mark ${escapeHtml(task.title)} complete"><div class="task-copy"><p class="task-title">${escapeHtml(task.title)}</p><div class="task-meta"><span class="badge">${task.frequency}</span>${task.done ? 'Completed' : 'Pending'}</div></div><button class="delete" data-delete="${task.id}" aria-label="Delete ${escapeHtml(task.title)}">Delete</button></article>`).join('');
  $('emptyState').hidden = filtered.length !== 0;
}

function escapeHtml(value) { const div = document.createElement('div'); div.textContent = value; return div.innerHTML; }
function completeTask(id) {
  const task = tasks.find(item => item.id === id); if (!task) return;
  task.done = !task.done;
  if (task.done) {
    const nextMonth = addMonths(task.month, task.frequency === 'Quarterly' ? 3 : 1);
    if (!tasks.some(item => item.title === task.title && item.month === nextMonth)) tasks.push({ id: crypto.randomUUID(), title: task.title, month: nextMonth, frequency: task.frequency, done: false });
  }
  save(); render();
}

$('addForm').addEventListener('submit', event => { event.preventDefault(); const title = $('taskTitle').value.trim(); if (!title) return; tasks.push({ id: crypto.randomUUID(), title, month: selectedMonth, frequency: $('frequency').value, done: false }); save(); $('taskTitle').value = ''; render(); });
$('taskList').addEventListener('click', event => { const checkbox = event.target.closest('.task-check'); const remove = event.target.closest('[data-delete]'); if (checkbox) completeTask(checkbox.dataset.id); if (remove) { tasks = tasks.filter(task => task.id !== remove.dataset.delete); save(); render(); } });
$('monthPicker').addEventListener('change', event => { selectedMonth = event.target.value; render(); });
$('previousMonth').addEventListener('click', () => { selectedMonth = addMonths(selectedMonth, -1); render(); });
$('nextMonth').addEventListener('click', () => { selectedMonth = addMonths(selectedMonth, 1); render(); });
$('todayButton').addEventListener('click', () => { selectedMonth = new Date().toISOString().slice(0, 7); render(); });
$('searchInput').addEventListener('input', render);
document.querySelectorAll('.filter').forEach(button => button.addEventListener('click', () => { activeFilter = button.dataset.filter; document.querySelectorAll('.filter').forEach(item => item.classList.toggle('active', item === button)); render(); }));
$('clearDone').addEventListener('click', () => { tasks = tasks.filter(task => !(task.month === selectedMonth && task.done)); save(); render(); });
$('exportExcel').addEventListener('click', async () => {
  const button = $('exportExcel');
  button.disabled = true;
  button.textContent = 'Saving...';
  try {
    const response = await fetch('/export-excel', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(tasks) });
    if (!response.ok) throw new Error('The local server did not accept the export.');
    const result = await response.json();
    alert(`${result.added} new task(s) appended to MyToDoList.xlsx in the project folder.`);
  } catch (error) {
    alert('Start the website with: python server.py\n\n' + error.message);
  } finally {
    button.disabled = false;
    button.textContent = 'Export Excel';
  }
});
$('themeButton').addEventListener('click', () => { document.body.classList.toggle('dark'); localStorage.setItem('monthly-todo-theme', document.body.classList.contains('dark') ? 'dark' : 'light'); });
$('resetData').addEventListener('click', () => { if (confirm('Reset the app to its demo task?')) { tasks = [...initialTasks]; save(); render(); } });
if (localStorage.getItem('monthly-todo-theme') === 'dark') document.body.classList.add('dark');
render();
