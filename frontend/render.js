const statusDiv = document.getElementById('status');
const logContent = document.getElementById('log-content');
const logContainer = document.getElementById('log-container');
const toggleLogSizeBtn = document.getElementById('toggle-log-size');

// Fonction pour basculer entre la vue normale et la vue plein écran des logs
toggleLogSizeBtn.addEventListener('click', function() {
    document.body.classList.toggle('fullscreen');
    
    if (document.body.classList.contains('fullscreen')) {
        toggleLogSizeBtn.innerText = "Réduire";
    } else {
        toggleLogSizeBtn.innerText = "Agrandir";
    }
});

// Fonction générique pour envoyer des requêtes à Flask
function sendRequest(url, method = 'POST') {
    fetch(`http://127.0.0.1:5000${url}`, { method: method })
        .then(response => response.json())
        .then(data => {
            statusDiv.innerText = data.status || 'Opération effectuée';
            statusDiv.classList.add('active');
        })
        .catch(error => {
            statusDiv.innerText = 'Une erreur s\'est produite';
            console.error('Erreur:', error);
        });
}

// Fonction pour récupérer les logs en temps réel
function fetchLogs() {
    fetch('http://127.0.0.1:5000/logs')
        .then(response => response.text())
        .then(data => {
            appendLog(data);  // Mise à jour des logs
        })
        .catch(error => {
            console.error('Erreur lors de la récupération des logs:', error);
        });
}

// Fonction pour ajouter des logs au conteneur et défiler automatiquement vers le bas
function appendLog(newLog) {
    const logContent = document.getElementById('log-content');
    logContent.textContent += newLog + "\n";
    
    // Défilement automatique vers le bas
    const logContentWrapper = document.getElementById('log-content-wrapper');
    logContentWrapper.scrollTop = logContentWrapper.scrollHeight;
}

// Fonction pour gérer les événements des boutons start/stop
function handleRpaButtonClick(button, method = 'POST') {
    const url = button.getAttribute('data-url');
    sendRequest(url, method);

    if (method === 'POST' && url.startsWith('/start')) {
        setInterval(fetchLogs, 3000);  // Récupérer les logs toutes les 3 secondes
    }
}

// Sélectionner les boutons start et stop pour tous les RPAs
const startButtons = document.querySelectorAll('[id^="start"]');
const stopButtons = document.querySelectorAll('[id^="stop"]');

// Gestion des clics pour démarrer les RPAs
startButtons.forEach(button => {
    button.addEventListener('click', () => handleRpaButtonClick(button));
});

// Gestion des clics pour arrêter les RPAs
stopButtons.forEach(button => {
    button.addEventListener('click', () => handleRpaButtonClick(button, 'POST'));
});
