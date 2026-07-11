const input=document.getElementById("pdf");
const drop=document.getElementById("dropZone");

input.addEventListener("change",()=>{

    if(input.files.length){

        document.getElementById("filename").innerHTML=input.files[0].name;

    }

});

drop.addEventListener("dragover",(e)=>{

    e.preventDefault();

    drop.classList.add("drag");

});

drop.addEventListener("dragleave",()=>{

    drop.classList.remove("drag");

});

drop.addEventListener("drop",(e)=>{

    e.preventDefault();

    drop.classList.remove("drag");

    input.files=e.dataTransfer.files;

    document.getElementById("filename").innerHTML=input.files[0].name;

});

async function convertir(){

    if(input.files.length===0){

        alert("Choisissez un PDF.");

        return;

    }

    const status=document.getElementById("status");

    status.innerHTML="⏳ Conversion en cours...";

    const data=new FormData();

    data.append("file",input.files[0]);

    const response=await fetch("https://attijaripdf2excel.onrender.com/convert",{

        method:"POST",

        body:data

    });

    if(!response.ok){

        status.innerHTML="❌ Une erreur est survenue.";

        return;

    }

    const blob=await response.blob();

    const url=window.URL.createObjectURL(blob);

    const a=document.createElement("a");

    a.href=url;

    a.download="releve.xlsx";

    a.click();

    status.innerHTML="✅ Conversion terminée.";

}
