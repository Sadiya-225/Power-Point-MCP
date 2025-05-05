document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('presentation-form');
    const statusDiv = document.getElementById('status');
    
    form.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const prompt = document.getElementById('prompt').value;
        
        if (!prompt) {
            showStatus('Please enter a prompt', 'error');
            return;
        }
        
        showStatus('Generating presentation...', 'info');
        
        // Create form data
        const formData = new FormData();
        formData.append('prompt', prompt);
        
        // Send the request
        fetch('/generate', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                showStatus(`Presentation generated successfully! ${data.details.slides_count} slides created.`, 'success');
            } else {
                showStatus(`Error: ${data.message}`, 'error');
            }
        })
        .catch(error => {
            showStatus(`Error: ${error.message}`, 'error');
        });
    });
    
    // Helper function to show status messages
    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = 'status-container';
        
        if (type === 'success') {
            statusDiv.classList.add('status-success');
        } else if (type === 'error') {
            statusDiv.classList.add('status-error');
        } else {
            statusDiv.style.display = 'block';
            statusDiv.style.backgroundColor = '#f0f0f0';
            statusDiv.style.border = '1px solid #ccc';
        }
    }
});