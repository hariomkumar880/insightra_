// Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-upload');
const fileNameDisplay = document.getElementById('file-name-display');

// Metrics Elements
const maxValEl = document.getElementById('max-value');
const minValEl = document.getElementById('min-value');
const profitEl = document.getElementById('total-profit');
const lossEl = document.getElementById('total-loss');

// Global Chart Instances
let barChartInstance = null;
let pieChartInstance = null;

// Initialize on Load
document.addEventListener('DOMContentLoaded', () => {
    const savedData = localStorage.getItem('insightra_data');
    if (savedData) {
        try {
            const parsed = JSON.parse(savedData);
            updateDashboard(parsed);
            fileNameDisplay.textContent = localStorage.getItem('insightra_filename') || 'Loaded from storage';
        } catch (e) {
            console.error("Error loading saved data", e);
        }
    }
});

// Drag and Drop Events
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
        handleFile(e.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFile(e.target.files[0]);
    }
});

function handleFile(file) {
    fileNameDisplay.textContent = file.name;
    
    // Check if PDF or Excel
    if (file.name.endsWith('.pdf')) {
        alert("PDF detected. Note: PDF extraction is limited in frontend. We recommend Excel (.xlsx) for accurate data charts.");
        // Basic mock processing for PDF to satisfy the request without a backend
        const mockData = generateMockData();
        localStorage.setItem('insightra_filename', file.name);
        updateDashboard(mockData);
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);
            
            if (json.length > 0) {
                const processed = processExcelData(json);
                localStorage.setItem('insightra_filename', file.name);
                updateDashboard(processed);
            } else {
                alert("The Excel file seems empty.");
            }
        } catch (error) {
            console.error(error);
            alert("Error parsing Excel file. Please ensure it's a valid .xlsx or .csv file.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function processExcelData(jsonArray) {
    // Attempt to dynamically find label columns and value columns
    let labels = [];
    let values = [];
    
    // Get keys
    const keys = Object.keys(jsonArray[0]);
    
    // Fallback if we can't find good columns
    let labelKey = keys[0]; 
    let valKey = keys.find(k => typeof jsonArray[0][k] === 'number') || keys[1];

    // Some common sense: if 'sales', 'profit', 'value' is in key name, use it for values
    const valueKeywords = ['sale', 'profit', 'loss', 'value', 'amount', 'price', 'total'];
    const possibleValKey = keys.find(k => valueKeywords.some(v => k.toLowerCase().includes(v)));
    if (possibleValKey && typeof jsonArray[0][possibleValKey] === 'number') {
        valKey = possibleValKey;
    }

    jsonArray.forEach(row => {
        labels.push(row[labelKey] || 'Unknown');
        // Parse float in case it's string number
        let v = parseFloat(row[valKey]);
        values.push(isNaN(v) ? 0 : v);
    });

    return { labels, values, datasetName: valKey };
}

function generateMockData() {
    return {
        labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
        values: [1200, 1900, 800, 2100, 1500],
        datasetName: 'PDF Extracted Data'
    };
}

function updateDashboard(data) {
    if (!data.values || data.values.length === 0) return;
    
    // Save to local storage
    localStorage.setItem('insightra_data', JSON.stringify(data));

    const max = Math.max(...data.values);
    const min = Math.min(...data.values);
    
    // Mocking Profit and Loss dynamically from the data for demonstration
    // If value > average, consider it profit-ish, else loss-ish just to show numbers
    const avg = data.values.reduce((a,b)=>a+b,0) / data.values.length;
    let profit = 0;
    let loss = 0;
    
    data.values.forEach(v => {
        if(v >= avg) profit += (v - avg);
        else loss += (avg - v);
    });

    // Format numbers
    maxValEl.textContent = formatNumber(max);
    minValEl.textContent = formatNumber(min);
    profitEl.textContent = '$' + formatNumber(profit);
    lossEl.textContent = '$' + formatNumber(loss);

    // Update Charts
    renderCharts(data.labels, data.values, data.datasetName);
}

function formatNumber(num) {
    return num.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function renderCharts(labels, values, datasetName) {
    // Common Chart Options for Dark Theme
    const chartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                labels: { color: '#f8fafc' }
            }
        },
        scales: {
            x: {
                ticks: { color: '#94a3b8' },
                grid: { color: 'rgba(255,255,255,0.1)' }
            },
            y: {
                ticks: { color: '#94a3b8' },
                grid: { color: 'rgba(255,255,255,0.1)' }
            }
        }
    };

    // Colors
    const bgColors = [
        'rgba(59, 130, 246, 0.7)',
        'rgba(16, 185, 129, 0.7)',
        'rgba(245, 158, 11, 0.7)',
        'rgba(239, 68, 68, 0.7)',
        'rgba(139, 92, 246, 0.7)',
        'rgba(236, 72, 153, 0.7)'
    ];

    // Bar Chart
    const barCtx = document.getElementById('barChart').getContext('2d');
    if (barChartInstance) barChartInstance.destroy();
    barChartInstance = new Chart(barCtx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: datasetName || 'Value',
                data: values,
                backgroundColor: 'rgba(59, 130, 246, 0.7)',
                borderColor: 'rgba(59, 130, 246, 1)',
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: chartOptions
    });

    // Pie Chart
    const pieCtx = document.getElementById('pieChart').getContext('2d');
    if (pieChartInstance) pieChartInstance.destroy();
    
    // Pie charts don't need XY scales
    const pieOptions = JSON.parse(JSON.stringify(chartOptions));
    delete pieOptions.scales;

    pieChartInstance = new Chart(pieCtx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: values,
                backgroundColor: bgColors,
                borderWidth: 1,
                borderColor: '#0f172a'
            }]
        },
        options: pieOptions
    });
}

// Chatbot Logic
const chatFab = document.getElementById('chat-fab');
const chatWindow = document.getElementById('chatbot-window');
const closeChat = document.getElementById('close-chat');
const chatInput = document.getElementById('chat-input');
const sendChat = document.getElementById('send-chat');
const chatMessages = document.getElementById('chat-messages');

chatFab.addEventListener('click', () => {
    chatWindow.classList.add('active');
    chatFab.style.display = 'none';
});

closeChat.addEventListener('click', () => {
    chatWindow.classList.remove('active');
    chatFab.style.display = 'block';
});

function handleChat() {
    const text = chatInput.value.trim();
    if (!text) return;

    // Add user message
    appendMessage(text, 'user-message');
    chatInput.value = '';

    // Bot logic
    setTimeout(() => {
        const lowerText = text.toLowerCase();
        let response = "I'm sorry, I don't understand that. Try asking about sales, profit, or loss.";

        if (lowerText.includes('sales') || lowerText.includes('sale')) {
            response = "Sales trend is increasing based on the current metrics.";
        } else if (lowerText.includes('profit')) {
            response = "Profit is positive. We are seeing a steady growth.";
        } else if (lowerText.includes('loss')) {
            response = "Loss is low, which is a great indicator of efficiency!";
        } else if (lowerText.includes('hello') || lowerText.includes('hi')) {
            response = "Hello! Upload your Excel data and ask me questions about it.";
        }

        appendMessage(response, 'bot-message');
    }, 500);
}

sendChat.addEventListener('click', handleChat);
chatInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') handleChat();
});

function appendMessage(text, className) {
    const div = document.createElement('div');
    div.className = `message ${className}`;
    div.textContent = text;
    chatMessages.appendChild(div);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}