// shared.js - Contains shared utility functions for HTML sidebars and dialogs

class UIManager {
    static setLoading(isLoading, progressElementId = 'progress-indicator', buttonElementId = 'extractButton') {
        const button = document.getElementById(buttonElementId);
        const progress = document.getElementById(progressElementId);

        if (button) {
            button.disabled = isLoading;
        }

        if (progress) {
            progress.classList.toggle('visible', isLoading);
        }
    }

    static showMessage(message, type = 'status', messageElementId = 'status-message') {
        const messageElement = document.getElementById(messageElementId);
        if (messageElement) {
            messageElement.textContent = message;
            messageElement.className = `message ${type} visible`;
        }
    }

    static clearMessages(messageElementId = 'status-message') {
        const messageElement = document.getElementById(messageElementId);
        if (messageElement) {
            messageElement.className = 'message';
            messageElement.textContent = '';
        }
         // Also clear the error message if it exists
        const errorElement = document.getElementById('error-message');
         if (errorElement) {
            errorElement.className = 'message';
            errorElement.textContent = '';
         }
    }
}

// Tooltip functions
function setupTooltips() {
    const tooltips = document.querySelectorAll('[data-tooltip]');
    // Check if a tooltip container already exists to avoid duplication
    let tooltipContainer = document.getElementById('tooltip-container');
    if (!tooltipContainer) {
        tooltipContainer = document.createElement('div');
        tooltipContainer.id = 'tooltip-container';
        tooltipContainer.className = 'tooltip-container';
        // Append to the body or a suitable common ancestor
        document.body.appendChild(tooltipContainer);
    }


    tooltips.forEach(element => {
        // Remove any existing tooltip event listeners to prevent duplicates
        // This is a bit tricky without storing references, but for this context,
        // re-adding should be fine if setupTooltips is called once on load.

        element.addEventListener('mouseenter', () => {
            // Check if tooltips are enabled (assuming a checkbox with id 'enable-tooltips' exists)
            const enableTooltipsCheckbox = document.getElementById('enable-tooltips');
            const tooltipsEnabled = enableTooltipsCheckbox ? enableTooltipsCheckbox.checked : true; // Default to true if checkbox not found

            if (tooltipsEnabled) {
                 const tooltipText = element.getAttribute('data-tooltip');
                 if (tooltipText) {
                    tooltipContainer.textContent = tooltipText;

                    // Position the tooltip near the element
                    const rect = element.getBoundingClientRect();
                    tooltipContainer.style.top = `${rect.bottom + 5}px`; // 5px below the element
                    tooltipContainer.style.left = `${rect.left + rect.width / 2}px`; // Center below the element
                    tooltipContainer.style.transform = 'translateX(-50%)'; // Center horizontally

                    tooltipContainer.classList.add('visible');
                 }
            }
        });

        element.addEventListener('mouseleave', () => {
            tooltipContainer.classList.remove('visible');
        });
         // Add focus/blur for accessibility
        element.addEventListener('focus', () => {
             const enableTooltipsCheckbox = document.getElementById('enable-tooltips');
             const tooltipsEnabled = enableTooltipsCheckbox ? enableTooltipsCheckbox.checked : true;

             if (tooltipsEnabled) {
                 const tooltipText = element.getAttribute('data-tooltip');
                 if (tooltipText) {
                     tooltipContainer.textContent = tooltipText;
                     const rect = element.getBoundingClientRect();
                     tooltipContainer.style.top = `${rect.bottom + 5}px`;
                     tooltipContainer.style.left = `${rect.left + rect.width / 2}px`;
                     tooltipContainer.style.transform = 'translateX(-50%)';
                     tooltipContainer.classList.add('visible');
                 }
             }
         });

         element.addEventListener('blur', () => {
             tooltipContainer.classList.remove('visible');
         });
    });
}

// Call setupTooltips on DOMContentLoaded in each HTML file after including shared.js