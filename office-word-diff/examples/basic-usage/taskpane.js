/* global Word, Office */
import { OfficeWordDiff } from '../../src/index.js';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initialize();
    } else {
        document.getElementById('result').innerHTML = 
            '<div class="result error">This add-in only works in Microsoft Word.</div>';
    }
});

function initialize() {
    const applyBtn = document.getElementById('applyDiff');
    const newTextArea = document.getElementById('newText');
    const resultDiv = document.getElementById('result');
    
    applyBtn.onclick = applyDiff;
    
    // Load current selection text into textarea
    loadSelection();
    
    async function loadSelection() {
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                
                if (selection.text && selection.text.trim()) {
                    newTextArea.value = selection.text;
                }
            });
        } catch (error) {
            console.error('Failed to load selection:', error);
        }
    }
    
    async function applyDiff() {
        const newText = newTextArea.value.trim();
        
        if (!newText) {
            showResult('Please enter new text to apply.', 'error');
            return;
        }
        
        applyBtn.disabled = true;
        applyBtn.textContent = 'Applying...';
        resultDiv.innerHTML = '';
        
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                
                if (!selection.text || !selection.text.trim()) {
                    throw new Error('Please select some text in the document first.');
                }
                
                const originalText = selection.text;
                
                // Create diff engine
                const differ = new OfficeWordDiff({
                    enableTracking: true,
                    logLevel: 'info',
                    onLog: (message, level) => {
                        console.log(`[${level}] ${message}`);
                    }
                });
                
                // Preview stats
                const stats = differ.getDiffStats(originalText, newText);
                console.log('Preview:', stats);
                
                // Apply the diff
                const result = await differ.applyDiff(context, selection, originalText, newText);
                
                if (result.success) {
                    const strategyNames = {
                        'token': 'Token Map (word-level)',
                        'sentence': 'Sentence Diff',
                        'block': 'Block Replace'
                    };
                    
                    showResult(
                        `✅ Successfully applied changes!<br>` +
                        `Strategy: ${strategyNames[result.strategyUsed] || result.strategyUsed}<br>` +
                        `Insertions: ${result.insertions}, Deletions: ${result.deletions}<br>` +
                        `Duration: ${result.duration}ms`,
                        'success'
                    );
                    
                    // Show preview stats
                    if (stats.totalChanges > 0) {
                        const statsHtml = `
                            <div class="stats">
                                Preview: ${stats.insertions} insertions, ${stats.deletions} deletions, 
                                ${stats.unchanged} unchanged tokens
                            </div>
                        `;
                        resultDiv.innerHTML += statsHtml;
                    }
                } else {
                    showResult(
                        `❌ Failed to apply changes. Check console for details.`,
                        'error'
                    );
                }
            });
            
        } catch (error) {
            showResult(`Error: ${error.message}`, 'error');
            console.error('Diff application error:', error);
        } finally {
            applyBtn.disabled = false;
            applyBtn.textContent = 'Apply Diff';
        }
    }
    
    function showResult(message, type) {
        resultDiv.innerHTML = `<div class="result ${type}">${message}</div>`;
    }
}
