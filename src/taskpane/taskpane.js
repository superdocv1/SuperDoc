
import { getAIContent, fetchSharePointListItems,sendForApproval } from "../service/api";

import PreviewIcon from './images/eye-open.png';
import ContentIcon from './images/content-icon.png';


let selectedText = '';
let promptText = '';
let generatedText = '';

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Word) {

      
        // Content Finder functionality
        const searchInput = document.querySelector('.custom-search');
        searchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                handleSearch(e.target.value);
            }
        });


        // Content Enhancer functionality
        document.getElementById('regenerateButton').onclick = handleRegenerate;
        document.getElementById('promptTextArea').oninput = handlePromptInput;
        document.getElementById('replace').onclick = handleReplace;
        document.getElementById('sendForApprovalButton').onclick = handleSendForApproval;

        // Listen for selection changes in the document
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, captureSelectedText);

        // Initial load of items for Content Finder
        
    }
});

async function showElementDlg(content,topic){
  var dialogUrl = 'https://'+location.host+'/dialog.html';
  const encodedTopic = encodeURIComponent(topic);
  const encodedContent = encodeURIComponent(content);
  dialogUrl += `?topic=${encodedTopic}&content=${encodedContent}`;
  Office.context.ui.displayDialogAsync(dialogUrl,{width:30,height:40,requireHTTPS:true,promptBeforeOpen: false},function(asyncResult){
    if(asyncResult.status != Office.AsyncResultStatus.Succeeded){

    }else{
      _dialog = asyncResult.value;
      _dialog.addEventListener(Microsoft.Office.WebExtension.EventType.DialogMessageReceived,messageHandlerElement);
    }
  })
}

function showInitialMessage() {
    const accordionContainer = document.getElementById('accordionExample');
    accordionContainer.innerHTML = '<p>Enter a search term and press Enter to find content.</p>';
}

async function handleSearch(searchTerm = '') {
    if (!searchTerm.trim()) {
        showInitialMessage();
        return;
    }

    try {
        const items = await fetchSharePointListItems(searchTerm);
        buildAccordion(items);
    } catch (error) {
        console.error('Error in search:', error);
        // Handle error (e.g., show error message to user)
        customAlert('An error occurred while searching. Please try again.');
    }
}

function buildAccordion(items) {
    console.log(items,"items")
    const accordionContainer = document.getElementById('accordionExample');
    accordionContainer.innerHTML = ''; // Clear existing items

    if (items.length === 0) {
        accordionContainer.innerHTML = '<p>No results found.</p>';
        return;
    }

    items.forEach((item, index) => {
        const accordionItem = createAccordionItem(item, index);
        accordionContainer.appendChild(accordionItem);
    });
}

function customAlert(message) {
    const alertDiv = document.createElement('div');
    alertDiv.style.cssText = `
        position: fixed;
        top: 10px;
        left: 50%;
        transform: translateX(-50%);
        background-color: #d4edda;
        color: #155724;
        padding: 10px 20px;
        border-radius: 5px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        z-index: 1000;
    `;
    alertDiv.textContent = message;

    document.body.appendChild(alertDiv);

    setTimeout(() => {
        alertDiv.remove();
    }, 3000); // Remove the alert after 3 seconds
}


function createAccordionItem(item, index) {
    // Function to truncate text to two lines
    function truncateToTwoLines(text, maxChars = 100) {
        if (text.length <= maxChars) return text;
        let truncated = text.substr(0, maxChars);
        let lastSpaceIndex = truncated.lastIndexOf(' ');
        if (lastSpaceIndex > -1) {
            truncated = truncated.substr(0, lastSpaceIndex);
        }
        return truncated + '...';
    }

    const truncatedContent = truncateToTwoLines(item.fields.TemplateContent || '');

    const accordionItem = document.createElement('div');
    accordionItem.className = 'accordion-item my-3';
    accordionItem.innerHTML = `
        <h2 class="accordion-header">
            <button class="accordion-button custom-accordian ${index !== 0 ? 'collapsed' : ''}" type="button" data-bs-toggle="collapse" data-bs-target="#collapse${index}" aria-expanded="${index === 0 ? 'true' : 'false'}" aria-controls="collapse${index}">
                <div class="d-flex flex-column">
                    <p class="font-10 font-medium black-v6 mb-1">${item.fields.TemplateTopic}</p>
                    <p class="mb-0 font-8 font-medium gray-v5">${item.lookupDetails?.fields?.TemplateName || ''}</p>
                </div>
            </button>
        </h2>
        <div id="collapse${index}" class="accordion-collapse collapse ${index === 0 ? 'show' : ''}" data-bs-parent="#accordionExample">
            <div class="accordion-body">
            <div>
                <p class="font-11 font-regular black-v6">${truncatedContent}</p>
                </div>
                <div class="d-flex gap-3">
                    <button class="white-btn font-10 font-medium color-black d-flex align-items-center gap-2 preview-btn">
                        <img src="${PreviewIcon}" alt="prev-icon" class="w-18">Content Preview
                    </button>
                    <button class="white-btn font-10 font-medium color-black d-flex align-items-center gap-2 insert-btn">
                        <img src="${ContentIcon}" alt="prev-icon" class="w-18">Insert Content
                    </button>
                </div>
            </div>
        </div>
    `;

    // Add event listeners for preview and insert buttons
    accordionItem.querySelector('.preview-btn').addEventListener('click', () => handlePreview(item.fields.TemplateTopic, item.fields.TemplateContent));
    accordionItem.querySelector('.insert-btn').addEventListener('click', () => handleInsert(item.fields.TemplateContent));

    return accordionItem;
}
async function captureSelectedText() {
    try {
        await Word.run(async (context) => {
            let selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            selectedText = selection.text;
        });
    } catch (error) {
        console.error('Error capturing selected text:', error);
    }
}

function handlePromptInput(event) {
    promptText = event.target.value;
}

async function handleRegenerate() {
  if (!selectedText) {
      customAlert('Please select some text in the document before regenerating.');
      return;
  }
  if (!promptText) {
      customAlert('Please enter a prompt before regenerating.');
      return;
  }

  try {
      generatedText = await getAIContent(selectedText, promptText);
      document.getElementById('responseTextArea').value = generatedText;
  } catch (error) {
      console.error('Error in text generation:', error);
      customAlert('An error occurred during the process. Please try again.');
  }
}


async function handleReplace() {
    if (!generatedText) {
        customAlert('Please regenerate text before replacing.');
        return;
    }

    try {
        await Word.run(async (context) => {
            let selection = context.document.getSelection();
            selection.insertText(generatedText, Word.InsertLocation.replace);
            await context.sync();
        });
        customAlert('Text successfully replaced!');
    } catch (error) {
        console.error('Error in text replacement:', error);
        customAlert('An error occurred during text replacement. Please try again.');
    }
}


// javascript
// function showPreviewDialog(content) {
//     // Create dialog container
//     const dialogContainer = document.createElement('div');
//     dialogContainer.className = 'dialog-container';
//     dialogContainer.style.cssText = `
//         position: fixed;
//         top: 0;
//         left: 0;
//         width: 100%;
//         height: 100%;
//         background-color: rgba(0, 0, 0, 0.5);
//         display: flex;
//         justify-content: center;
//         align-items: center;
//         z-index: 1000;
//     `;

//     // Create dialog content
//     const dialog = document.createElement('div');
//     dialog.className = 'dialog';
//     dialog.style.cssText = `
//         background-color: white;
//         padding: 20px;
//         border-radius: 5px;
//         max-width: 80%;
//         max-height: 80%;
//         overflow-y: auto;
//     `;

//     // Add content to dialog
//     dialog.innerHTML = `
//         <h2>Content Preview</h2>
//         <p>${content}</p>
//         <button id="closeDialog" class="btn primary-btn font-14 font-medium">Close</button>
//     `;

//     // Add dialog to container
//     dialogContainer.appendChild(dialog);

//     // Add container to body
//     document.body.appendChild(dialogContainer);

//     // Add event listener to close button
//     document.getElementById('closeDialog').addEventListener('click', () => {
//         document.body.removeChild(dialogContainer);
//     });

//     // Close dialog when clicking outside
//     dialogContainer.addEventListener('click', (e) => {
//         if (e.target === dialogContainer) {
//             document.body.removeChild(dialogContainer);
//         }
//     });
// }



function handlePreview(content,topic) {
    console.log("Preview content:", content); // For debugging
    showElementDlg(content,topic);
}

// Set up event listener for preview buttons
document.addEventListener('click', (e) => {
    if (e.target.closest('.preview-btn')) {
        const content = e.target.closest('.accordion-item').querySelector('.accordion-body p').innerHTML;
        handlePreview(content);
    }
});



// Make sure to call this function when your add-in loads
function initializeAddIn() {
    // Other initialization code...

    // Set up event listener for preview buttons
    document.addEventListener('click', (e) => {
        if (e.target.closest('.preview-btn')) {
            const content = e.target.closest('.accordion-item').querySelector('.accordion-body p').textContent;
            const topic = e.target.closest('.accordion-item').querySelector('.accordion-header .font-10').textContent;
          
            handlePreview(content,topic);
        }
    });
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initializeAddIn();
    }
});


async function handleInsert(content) {
  try {
      // Strip HTML tags and decode HTML entities
      const plainTextContent = decodeHTMLEntities(stripHTMLTags(content));

      await Word.run(async (context) => {
          let selection = context.document.getSelection();
          selection.insertText(plainTextContent, Word.InsertLocation.replace);
          await context.sync();
      });
      customAlert('Content successfully inserted!');
  } catch (error) {
      console.error('Error inserting content:', error);
      customAlert('An error occurred while inserting the content. Please try again.');
  }
}

// Function to strip HTML tags


// Function to decode HTML entities
function decodeHTMLEntities(text) {
    const textArea = document.createElement('textarea');
    textArea.innerHTML = text;
    return textArea.value;
}
// Initial search function call removed from here

function stripHTMLTags(html) {
    const tmp = document.createElement('DIV');
    tmp.innerHTML = html;
    return tmp.textContent || tmp.innerText || '';
}

async function handleSendForApproval() {
    try {
        // Get the document URL
        const documentUrl = await getDocumentUrl();

        console.log(documentUrl)
        // Extract document name from URL
        const documentName = getDocumentNameFromUrl(documentUrl);

        // Send for approval
        const response = await sendForApproval(documentName);

        if (response) {
            showApprovalSentPopup();
        } else {
            throw new Error('Approval request failed');
        }
    } catch (error) {
        console.error('Error in sending for approval:', error);
        customAlert('An error occurred while sending for approval. Please try again.');
    }
}

async function getDocumentUrl() {
    return new Promise((resolve, reject) => {
        Office.context.document.getFilePropertiesAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value.url);
            } else {
                reject(new Error('Failed to get document URL'));
            }
        });
    });
}

function getDocumentNameFromUrl(url) {
    const urlParts = url.split('/');
    return urlParts[urlParts.length - 1];
}

function showApprovalSentPopup() {
    const popupDiv = document.createElement('div');
    popupDiv.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background-color: #d4edda;
        color: #155724;
        padding: 20px;
        border-radius: 5px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        z-index: 1000;
    `;
    popupDiv.textContent = 'Your document has been sent for approval.';

    document.body.appendChild(popupDiv);

    setTimeout(() => {
        popupDiv.remove();
    }, 3000); // Remove the popup after 3 seconds
}

   

