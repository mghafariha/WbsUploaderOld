<script>
$(document).ready(function () {
    cascadeLookup({
        listId: '0097C3C3-4302-4873-A3FB-E9F5194520A1', //Guid of List
        parentField: "MainOperation",//fieldParent in Destination list
        childField: "Title",//fieldchild in Destination list
        masterField: "MainOperation",//master Field in current List
        detailField: "DefectTitle",//detail Field in current List
    });

    window.PreSaveAction = function PreSaveAction() {
        // add other your codes
        var x = $("#ClientFormPostBackValue_cc4b55ab-56a9-447e-b56c-af264c74fd3a_NumberRepetitions").val();
        var x2 = parseInt(x);
    
        var y = $("#ClientFormPostBackValue_cc4b55ab-56a9-447e-b56c-af264c74fd3a_NumberRepetitionsInContract").val();
        var y2 = parseInt(y);
        if (y2 > x2) {
            alert("تعداد تکرار در چارچوب پیمان باید کمتر از تعداد تکرار باشد ");
            return false; 
        }

        return true; 
        // use return true; if confirm or return false...
    }
});
</script>