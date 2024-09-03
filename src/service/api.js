import axios from 'axios';

export const callApi = async(payload ) => {
    const url = "https://ava-eus-m365-function-teamsbot.azurewebsites.net/api/superDocApi?";
    // const url = "http://localhost:7071/api/superDocApi"
    payload = {...payload, token : "U2FsdGVkX193A1/d5YIeSo1XKuGepAVPccLJUDfWcy8=", userId: "123456"}
    const headers = { 
        'Content-Type': 'application/json',
      }
    const response = await axios.post(url, payload, {headers : headers});
    return {status : response.status, data : response.data};
}

import CryptoJS from 'crypto-js';
// Commands to install crypto-js : 1. npm i crypto-js  2.  npm i --save-dev @types/crypto-js


const secretKey = "6f5206413a3c3cacf4fb53518abd1ca1e63b57328b61217e6003222946f6a5fe";

export const encrypt = (data) => {
  const token = CryptoJS.AES.encrypt(data, secretKey).toString();
  console.log("JWT token ", token);
  return token;
}

export const decrypt = (encryptedData) => {
  const bytes = CryptoJS.AES.decrypt(encryptedData, secretKey);
  const decrytedData = bytes.toString(CryptoJS.enc.Utf8);
  return decrytedData;
}

const getAccessToken = async() => {
  const accessToken = await callApi({ type: "getaccesstoken" });
  return accessToken.data;
}

export async function getJWTtoken() {
 
  try {
      const AiUrl = `https://superdoc-dev.superdesk.ai/create_token`;
      // const userId = stringify(context.activity.from.aadObjectId);
      const userId = JSON.stringify({ "userId": "user123" });
      const headers = { 'Content-Type': 'application/json' };
      const resp= await axios.post(AiUrl, userId, { headers: headers });
      console.log("JWT token:", resp.data.token);

      return resp.data.token;


  }
  catch (e) {
      console.error('Access Token Error', e);
  }
}
export async function getAIContent(selectedtext,userprompt) {

    // try {
    //     //  P1_PC_10
        
    //     var AzureOpenAIUrl =
    //       "https://ava-si-cdgen-opai.openai.azure.com/openai/deployments/OpenAI-GEPT-4/chat/completions?api-version=2024-02-15-preview";
    //     var body = {
    //       messages: [
    //         {
    //           role: "system",
    //           content: "your objective is to rephrase the Content based on the users wish like in what way that the user is asking",
    //         },
    //         {
    //           role: "user",
    //           content: selectedtext+userprompt,
    //         },
    //       ],
    //       temperature: 0.7,
    //       top_p: 0.95,
    //       frequency_penalty: 0,
    //       presence_penalty: 0,
    //       max_tokens: 800,
    //       stop: null,
    //     };
    //     var headers = { "api-key": "f456da61f9204a808dedd6bf96f5f0bd", "Content-Type": "application/json" };
    //     var resp = await axios.post(AzureOpenAIUrl, body, { headers: headers });
       
    //     console.log(resp)
    //     console.log(resp.data.choices[0].message.content)
    //     console.log(resp.status)
        
    //     // console.log(resp.data.choices[0])
    //     // const jsonContent = content.substring(8, content.length - 3);
    //     // // Parse the JSON content
    //     // const parsedContent = JSON.parse(jsonContent);
    
    //     // Extract the chat summary
    //     // const chatSummary = parsedContent.chat_summary;
    //     return resp.data.choices[0].message.content
    //     // eslint-disable-next-line no-unreachable
        
    //   } catch (e) {
    //     console.log(e);
    //   }

    try {
      const JWTtoken = await getJWTtoken();

      console.log("JWTtoken", JWTtoken);
      const getAIURL = `https://superdoc-dev.superdesk.ai/rewrite`

      const data = JSON.stringify({
         "SelectedContent":selectedtext,
          "Token": JWTtoken,
          "userId": "user123",
          "UserPrompt":userprompt
         
      })
      const headers = {
          'Authorization': `Bearer ${JWTtoken}`,
          'Content-Type': 'application/json'
      };

      const response = await axios.post(getAIURL, data, { headers })
      console.log(response.data.rewritten_sentence);

      
      return response.data.rewritten_sentence
  } catch (e) {
      console.error('Access Token Error', e);
  }
}



  

export async function fetchSharePointListItems(searchTerm) {
    try {
        let access_token= await getAccessToken();
        const siteId = '1384ebf8-ad08-4cde-9c19-962dcd0caec5';
        const listId = '45d9333a-1619-4b4d-94a2-1bd55a1bba94';
        const apiUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields($select=*)&$filter=startswith((fields/TemplateTopic), '${encodeURIComponent(searchTerm)}')`;
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${access_token}`,
            'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly'
        };

        const resp = await axios.get(apiUrl, { headers: headers });
        const items = resp.data.value;

        // Separate items with TemplateId and those without
        const itemsWithTemplateId = items.filter(item => item.fields.TemplateId);
        const itemsWithoutTemplateId = items.filter(item => !item.fields.TemplateId);

        // Fetch details for items with TemplateId
        const lookupDetailsPromises = itemsWithTemplateId.map(async (item) => {
            const lookupUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/99cc487a-1451-476c-a1ff-69d23fc5f97b/items/${item.fields.TemplateId}?$expand=fields($select=*)`;
            try {
                const lookupResponse = await axios.get(lookupUrl, { headers: headers });
                return {
                    ...item,
                    lookupDetails: lookupResponse.data
                };
            } catch (error) {
                console.error(`Error fetching details for TemplateId ${item.fields.TemplateId}:`, error);
                return item; // Return the original item if lookup fails
            }
        });

        const itemsWithLookupDetails = await Promise.all(lookupDetailsPromises);

        // Combine items with lookup details and items without TemplateId
        const combinedItems = [...itemsWithLookupDetails, ...itemsWithoutTemplateId];

        console.log("Combined Items:", combinedItems);
        return combinedItems;

    } catch (e) {
        console.error('Error details:', e.response ? e.response.data : e.message);
        throw e; // Re-throw the error for the caller to handle
    }
}



export async function sendForApproval(documentName) {
    try {
      
      const apiUrl = `https://prod-28.westus.logic.azure.com:443/workflows/e65b8cb4c36244a68825677e1d3f473f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=UoezSGRHse6oJbHq_Z-kAgUpVdGtnprHuEZlus1A_tk`;
  
      
  
      const data = {
        filename: documentName
      };
  
      const response = await axios.post(apiUrl, data);
  
      return response.status;
    } catch (error) {
      console.error('Error in sendForApproval:', error);
      throw error;
    }
  }