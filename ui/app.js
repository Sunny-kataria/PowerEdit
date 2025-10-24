function send(msg) { window.chrome?.webview?.postMessage(msg); }
const statusEl = document.getElementById('status');
const toastEl = document.getElementById('toast');
const runBtn = document.getElementById('run');
const actionEl = document.getElementById('action');

function setStatus(t) { statusEl.textContent = t; }
function toast(t) {
    toastEl.textContent = t; toastEl.hidden = false;
    toastEl.classList.add('show');
    setTimeout(() => toastEl.classList.remove('show'), 1800);
}

document.getElementById('browse').addEventListener('click', () => {
    setStatus('Opening file dialog...');
    send({ type: 'browse' });
});

actionEl.addEventListener('change', () => {
    runBtn.disabled = !actionEl.value;
});

runBtn.addEventListener('click', () => {
    const v = actionEl.value;
    if (!v) return;
    setStatus('Running: ' + v + ' ...');
    send({ type: 'runAction', action: v });
});

document.getElementById('togglePin').addEventListener('click', () => {
    send({ type: 'togglePin' });
    toast('Pin toggle (not implemented yet)');
});

window.chrome?.webview?.addEventListener('message', (e) => {
    const msg = e.data;
    if (msg?.type === 'fileChosen') {
        document.getElementById('fileName').textContent = msg.name || '(unnamed)';
        setStatus('Loaded: ' + (msg.name || 'file'));
        toast('File loaded');
    }
    if (msg?.type === 'notify' || msg?.type === 'toast') {
        setStatus(msg.text || 'Done.');
        toast(msg.text || 'Done');
    }
});
