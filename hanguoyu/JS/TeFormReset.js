function FormReset(fid){
    var FormObject;
    FormObject = document.getElementById(fid);
    if(FormObject){
        FormObject.innerHTML = FormObject.innerHTML;
    }
}