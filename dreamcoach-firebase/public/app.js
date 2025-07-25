import { initializeApp } from 'https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js';
import { getAnalytics } from 'https://www.gstatic.com/firebasejs/9.23.0/firebase-analytics.js';

const firebaseConfig = {
  apiKey: "AIzaSyBUmGhmZkktNU-_QwFtERl-T6oRf-_sAgQ",
  authDomain: "dreamanalysis-39322.firebaseapp.com",
  projectId: "dreamanalysis-39322",
  storageBucket: "dreamanalysis-39322.firebasestorage.app",
  messagingSenderId: "222523234712",
  appId: "1:222523234712:web:b508d345729d98ad4133ed",
  measurementId: "G-0RKJVBCC8K"
};

const app = initializeApp(firebaseConfig);
getAnalytics(app);

const form = document.getElementById('dreamForm');
const resultDiv = document.getElementById('result');

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  resultDiv.classList.add('hidden');
  const data = {
    name: document.getElementById('name').value,
    email: document.getElementById('email').value,
    dreamDate: document.getElementById('dreamDate').value,
    description: document.getElementById('description').value
  };

  const resp = await fetch('/submitDream', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  });
  const json = await resp.json();
  if (json.success) {
    resultDiv.textContent = 'Analysis:\n' + json.analysis;
    resultDiv.classList.remove('hidden');
  } else {
    resultDiv.textContent = 'Error: ' + json.error;
    resultDiv.classList.remove('hidden');
  }
});
