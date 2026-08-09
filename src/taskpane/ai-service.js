// ======================================
// Universal AI Service
// Supports multiple providers
// ======================================



async function askAI(prompt){



    const settings =

    JSON.parse(

        localStorage.getItem("AI_SETTINGS")

    );




    if(!settings){


        return "⚠ لم يتم إعداد مزود الذكاء الاصطناعي";


    }




    if(!settings.key){


        return "⚠ لم يتم إدخال مفتاح API";


    }






    switch(settings.provider){



        case "openai":


            return await callOpenAI(

                prompt,

                settings

            );





        case "gemini":


            return await callGemini(

                prompt,

                settings

            );





        case "openrouter":


            return await callOpenRouter(

                prompt,

                settings

            );





        case "anthropic":


            return await callAnthropic(

                prompt,

                settings

            );





        case "groq":


            return await callGroq(

                prompt,

                settings

            );





        default:


            return await callCustom(

                prompt,

                settings

            );



    }



}









// ======================================
// OpenAI
// ======================================


async function callOpenAI(prompt,s){



let url =

s.url ||

"https://api.openai.com/v1/chat/completions";




let response = await fetch(url,{

method:"POST",


headers:{


"Authorization":

"Bearer "+s.key,


"Content-Type":

"application/json"


},


body:JSON.stringify({


model:s.model,


messages:[

{

role:"user",

content:prompt

}

]


})


});




let data =
await response.json();



return extractOpenAI(data);



}








// ======================================
// OpenRouter
// ======================================


async function callOpenRouter(prompt,s){



let url =

s.url ||

"https://openrouter.ai/api/v1/chat/completions";




let response = await fetch(url,{


method:"POST",


headers:{


"Authorization":

"Bearer "+s.key,


"Content-Type":

"application/json"


},


body:JSON.stringify({


model:s.model,


messages:[

{

role:"user",

content:prompt

}

]


})


});




let data =
await response.json();



return extractOpenAI(data);



}








// ======================================
// Gemini
// ======================================


async function callGemini(prompt,s){



let url =

(s.url ||

"https://generativelanguage.googleapis.com/v1beta/models/"
+s.model+
":generateContent")

+"?key="+s.key;





let response =
await fetch(url,{


method:"POST",


headers:{


"Content-Type":

"application/json"


},


body:JSON.stringify({


contents:[

{


parts:[

{

text:prompt

}

]


}

]


})


});





let data =
await response.json();





try{


return data

.candidates[0]

.content

.parts[0]

.text;



}

catch(e){


return JSON.stringify(data);


}



}









// ======================================
// Anthropic Claude
// ======================================


async function callAnthropic(prompt,s){



return "تم اختيار Anthropic - سيتم تفعيل الاتصال لاحقًا";



}









// ======================================
// Groq
// ======================================


async function callGroq(prompt,s){



return "تم اختيار Groq - سيتم تفعيل الاتصال لاحقًا";



}









// ======================================
// Custom Provider
// ======================================


async function callCustom(prompt,s){



return "تم اختيار مزود مخصص";



}









function extractOpenAI(data){



try{


return data

.choices[0]

.message

.content;


}

catch(e){


return JSON.stringify(data);


}



}