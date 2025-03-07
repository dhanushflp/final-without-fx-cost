document.getElementById('processingForm').addEventListener('submit', function(e) {
    e.preventDefault();
    
    const form = e.target;
    const formData = new FormData(form);
    
    // Show loading spinner
    document.getElementById('loadingSpinner').style.display = 'block';
    const resultArea = document.getElementById('resultArea');
    resultArea.innerHTML = '';

    fetch('/process', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        // Hide loading spinner
        document.getElementById('loadingSpinner').style.display = 'none';

        if (data.success) {
            resultArea.innerHTML = `
                <div class="success">
                    <h3>Processing Completed!</h3>
                    <p>Output file: ${data.finalFile}</p>
                    <a href="/download/${encodeURIComponent(data.finalFile)}" class="download-link">
                        Download Output File
                    </a>
                </div>
            `;
        } else {
            resultArea.innerHTML = `
                <div class="error">
                    <h3>Processing Failed</h3>
                    <p>${data.error}</p>
                </div>
            `;
        }
    })
    .catch(error => {
        // Hide loading spinner
        document.getElementById('loadingSpinner').style.display = 'none';
        
        resultArea.innerHTML = `
            <div class="error">
                <h3>Network Error</h3>
                <p>${error}</p>
            </div>
        `;
    });
});