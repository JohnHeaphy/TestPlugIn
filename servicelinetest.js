function Form_Onload()
{
//Rag_Status('OnLoad');
var SR = Xrm.Page.getAttribute("statuscode").getValue();
var FT = Xrm.Page.ui.getFormType();


    // Create form = True //
    if (FT == 1 && Xrm.Page.getAttribute("ccx_draft").getValue()== 0)
    {
    alert ("WARNING\n\nYou may only add Service Lines to a Draft Contract.\n\nThis form will now close.");
    Xrm.Page.ui.close();
    }

    // Status = Draft //
    if (SR == 1)
    {
    Check_Visability("OnLoad");
    Surplus_Deficit ("OnLoad");
    Short_Term();
    FunderLineValue('OnLoad');
    Comma_In_Service_Type();

    }

    // Status != Draft //
    if (SR != 1)
    {
//    alert (" StatusCode = " + Xrm.Page.getAttribute("statuscode").getValue())
    formdisable(true);
    Surplus_Deficit ("OnLoad");
    }
}



function Check_Visability(Trigger)
{
Xrm.Page.ui.controls.get("ccx_service_line_checked").setDisabled(true);

var SR = Xrm.Page.data.entity.attributes.get("statuscode").getValue();
var FT = Xrm.Page.ui.getFormType();
var SF = Xrm.Page.data.entity.attributes.get("ccx_csservicefundedid").getValue();
var T1 = Xrm.Page.data.entity.attributes.get("ccx_t1code").getValue();
var T2 = Xrm.Page.data.entity.attributes.get("ccx_t2code").getValue();
var TV = Xrm.Page.data.entity.attributes.get("ccx_totalcost").getValue();
var FV = Xrm.Page.data.entity.attributes.get("ccx_funder_line_value").getValue();
var BA = Xrm.Page.data.entity.attributes.get("ccx_billed_in_arrears").getValue();
var ST = Xrm.Page.data.entity.attributes.get("ccx_short_term_contract").getValue();
//var CT = Xrm.Page.data.entity.attributes.get("ccx_contract_type_code").getValue();

//alert ( " Trigger = " + Trigger +
//        "\n Form Type = " + FT +
//        "\n T1 Code = " + T1 +
//        "\n T2 Code = " + T2 +
//        "\n Total Value = " + TV +
//        "\n Service funded = " + SF +
//        "\n Status Reason = " + SR +
//        "\n Funder Value = " + FV +
//        "\n Billed in Arrears = " + BA +
//        "\n Short term contract = " + ST);

    if (FT != 1 && (T2 != null)&& TV >0 && SF != null && SR == 1 && (FV >0 || BA == true || ST == true))
    {
//    alert ("Check = True");
    Xrm.Page.ui.controls.get("ccx_service_line_checked").setDisabled(false);
    }
    else
    {
//    alert ("Check = False");
    }
}

function Contract_Expenditure(Trigger)
{
var TSC = 0.00;
CT = Xrm.Page.data.entity.attributes.get("ccx_calculate_totals").getValue();

//alert ("CT = " + CT + " TSC = " + " Trigger = " + Trigger);
    if (Trigger == 'OnChange'&& CT == true)
    {
        TSC =
    Xrm.Page.data.entity.attributes.get("ccx_salaries").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_employersnipension").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_trainingconferences").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_humanresourcessupport").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_carmileage").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_othercarexpenses" ).getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_mealsaccommodation").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_othertravelexpenses").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_totalstaffcosts").setValue(TSC);
    Xrm.Page.data.entity.attributes.get("ccx_totalstaffcosts").setSubmitMode("always");
    Total_Cost();
    }
}

function Direct_Expenditure(Trigger)
{
var TODE =0.00;
CT = Xrm.Page.data.entity.attributes.get("ccx_calculate_totals").getValue();

    if (Trigger == 'OnChange'&& CT == true)
    {
    TODE =
    Xrm.Page.data.entity.attributes.get("ccx_postage").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_telephone").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_catering").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_stationery").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_photocopying").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_itequipmentsupport").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_professionalfees").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_premisescosts").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_strokeclubs").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_publications").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_advertising").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_clientmaterialsmiscellaneous").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_volunteercosts").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_patienttransport").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_roomhire" ).getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_groupmaterials" ).getValue();
    Xrm.Page.data.entity.attributes.get("ccx_totalotherdirectexpenditure").setValue(TODE);
    Xrm.Page.data.entity.attributes.get("ccx_totalotherdirectexpenditure").setSubmitMode("always");
    Total_Cost();
    }
}

function Managment_Support_Cost(Trigger)
{
var MSC =0.00;
CT = Xrm.Page.data.entity.attributes.get("ccx_calculate_totals").getValue();

    if (Trigger == 'OnChange'&& CT == true)
    {
    MSC =
    Xrm.Page.data.entity.attributes.get("ccx_directsupervision").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_regionalmanagement").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_centraloperationsmanagement").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_facilitiessupporthealthandsafety").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_qualityassuranceandcentralsupport").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_totalmanagementandsupportcosts").setValue(MSC);
    Xrm.Page.data.entity.attributes.get("ccx_totalmanagementandsupportcosts").setSubmitMode("always");
    Total_Cost();
    }
}

function Total_Other_Cost(Trigger)
{
var TOC =0.00;
CT = Xrm.Page.data.entity.attributes.get("ccx_calculate_totals").getValue();

    if (Trigger == 'OnChange'&& CT == true)
    {
    TOC =
    Xrm.Page.data.entity.attributes.get("ccx_short_term_contract_cost").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_my_stroke_guide_cost").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_irrecoverablevat").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_total_other_cost").setValue(TOC);
    Xrm.Page.data.entity.attributes.get("ccx_total_other_cost").setSubmitMode("always");
    Total_Cost();
    }
 }


function Total_Cost(Trigger)
{
var TDSC =0.00; // Total direct and staff cost
var TC =0.00;   // Total Cost
var TMSC =0.00  // Total managment suport and other costs
CT = Xrm.Page.data.entity.attributes.get("ccx_calculate_totals").getValue();

    if (CT == true)
    {
    TDSC=
    Xrm.Page.data.entity.attributes.get("ccx_totalstaffcosts").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_totalotherdirectexpenditure").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_totaldirectandstaffcost").setValue(TDSC);
    Xrm.Page.data.entity.attributes.get("ccx_totaldirectandstaffcost").setSubmitMode("always");
    //
    TMSC=
    Xrm.Page.data.entity.attributes.get("ccx_totalmanagementandsupportcosts").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_total_other_cost").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_totalmanagementsupportandothercost").setValue(TMSC);
    Xrm.Page.data.entity.attributes.get("ccx_totalmanagementsupportandothercost").setSubmitMode("always");
    //
    TC=
    Xrm.Page.data.entity.attributes.get("ccx_totaldirectandstaffcost").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_totalmanagementsupportandothercost").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_managementchargenotpreviouslyincluded").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_totalcost").setValue(TC);
    Xrm.Page.data.entity.attributes.get("ccx_totalcost").setSubmitMode("always");
    Surplus_Deficit("OnChange");
    }
}

function FunderLineValue(Trigger){
var FormType = Xrm.Page.ui.getFormType();
if (FormType != 1){

//alert ("FunderLineValue");
var ServiceLineID = Xrm.Page.data.entity.getId();
var FT = Xrm.Page.ui.getFormType();
var sum = 0.00;

var req = new XMLHttpRequest();
req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/Ccx_csfunderlineSet?$select=ccx_original_invoice_value&$filter=ccx_csservicelineid/Id eq (guid'"+ServiceLineID+"') and statuscode/Value eq 9", false);
req.setRequestHeader("Accept", "application/json");
req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
req.onreadystatechange = function () {
    if (this.readyState === 4) {
        this.onreadystatechange = null;
        if (this.status === 200) {
            var returned = JSON.parse(this.responseText).d;
            var results = returned.results;
            for (var i = 0; i < results.length; i++) {
                var ccx_original_invoice_value = results[i].ccx_original_invoice_value;
                var LineValue = parseFloat(ccx_original_invoice_value.Value);
//                alert("LineValue = " + LineValue);
                sum = sum + LineValue;
            }
        }
        else {
            alert(this.statusText);
        }
    }
};
req.send();

//alert ("Funder Line Value = " + sum);

Xrm.Page.data.entity.attributes.get("ccx_funder_line_value").setValue(sum);
Xrm.Page.data.entity.attributes.get("ccx_funder_line_value").setSubmitMode("always");

// Start Funtions
Surplus_Deficit("OnChange");

//var VolunterGrid = document.getElementById("Funder_Line");
//VolunterGrid.control.add_onRefresh(FunderLineValue);
}
}

 function Surplus_Deficit(Trigger)
 {
//alert ("Surplus Deficit Started");
 var FSD=0.00;
 var DSD=0.00;
 CT = Xrm.Page.data.entity.attributes.get("ccx_calculate_totals").getValue();
    if (CT == true && Trigger == 'OnChange')
    {
    DSD=
    Xrm.Page.data.entity.attributes.get("ccx_funder_line_value").getValue()-
    Xrm.Page.data.entity.attributes.get("ccx_totalcost").getValue()+
    Xrm.Page.data.entity.attributes.get("ccx_managementchargenotpreviouslyincluded").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_directsurplusordeficit").setValue(DSD);
    Xrm.Page.data.entity.attributes.get("ccx_directsurplusordeficit").setSubmitMode("always");
    FSD=
    Xrm.Page.data.entity.attributes.get("ccx_funder_line_value").getValue()-
    Xrm.Page.data.entity.attributes.get("ccx_totalcost").getValue();
    Xrm.Page.data.entity.attributes.get("ccx_finalsurplusordeficit").setValue(FSD);
    Xrm.Page.data.entity.attributes.get("ccx_finalsurplusordeficit").setSubmitMode("always");
    }
    if (Xrm.Page.data.entity.attributes.get("ccx_finalsurplusordeficit").getValue() > 0)
    {
//    document.getElementById("ccx_finalsurplusordeficit_d").style.backgroundColor = '#99FF00';
//    document.getElementById("ccx_finalsurplusordeficit_c").style.backgroundColor = '#99FF00';
    }
    else
    {
//    document.getElementById("ccx_finalsurplusordeficit_d").style.backgroundColor = '#FF3300';
//    document.getElementById("ccx_finalsurplusordeficit_c").style.backgroundColor = '#FF3300';
    }
    if (Xrm.Page.data.entity.attributes.get("ccx_directsurplusordeficit").getValue() > 0)
    {
//    document.getElementById("ccx_directsurplusordeficit_d").style.backgroundColor = '#99FF00';
//    document.getElementById("ccx_directsurplusordeficit_c").style.backgroundColor = '#99FF00';
    }
    else
    {
//    document.getElementById("ccx_directsurplusordeficit_d").style.backgroundColor = '#FF3300';
//    document.getElementById("ccx_directsurplusordeficit_c").style.backgroundColor = '#FF3300';
    }

Check_Visability("Surplus_Deficit");
}

function Checked(Trigger)
{
var UserName = Xrm.Page.context.getUserName();
var CDate = new Date();
var SLC = Xrm.Page.data.entity.attributes.get("ccx_service_line_checked").getValue();
//alert ("FLC = " + FLC + " UserName " + UserName);

    if (SLC == 2 && Trigger == 'OnChange')
    {
    var result = confirm("Warning!\n\nChecking this Service Line will Save and Lock the record\n\nDo you wish to continue?  ");

        if (result)
        {
//        alert("You pressed OK!");
        Xrm.Page.data.entity.attributes.get("ccx_service_line_checked_name").setValue(UserName);
        Xrm.Page.data.entity.attributes.get("ccx_service_line_checked_name").setSubmitMode("always");
        Xrm.Page.data.entity.attributes.get("ccx_service_line_checked_date").setValue(CDate);
        Xrm.Page.data.entity.attributes.get("ccx_service_line_checked_date").setSubmitMode("always");
        Xrm.Page.data.entity.attributes.get("statuscode").setValue(3);
        Xrm.Page.data.entity.attributes.get("statuscode").setSubmitMode("always");
        Xrm.Page.data.entity.save();
        formdisable(true);
        }
        else
        {
//        alert("You pressed Cancel!");
        Xrm.Page.data.entity.attributes.get("ccx_service_line_checked").setValue(1);
        }
    }
}

function formdisable(disablestatus)
{
    var allAttributes = Xrm.Page.data.entity.attributes.get();
    for (var i in allAttributes) {
           var myattribute = Xrm.Page.data.entity.attributes.get(allAttributes[i].getName());
           var myname = myattribute.getName();
           if (Xrm.Page.getControl(myname) != null) Xrm.Page.getControl(myname).setDisabled(disablestatus);
    }
 Xrm.Page.ui.controls.get("ccx_overallragstatus").setDisabled(false);
 Xrm.Page.ui.controls.get("ccx_regionalmanagersassessment" ).setDisabled(false);
} // formdisable


function Short_Term()
{
var ST = Xrm.Page.data.entity.attributes.get("ccx_short_term_contract").getValue();
    if (ST == false)
    {
      Xrm.Page.ui.controls.get("ccx_short_term_contract_cost").setDisabled(true);
    }
}

function Rag_Status(Trigger){
//alert ("Rag Started\nTrigger = " + Trigger);
var Rag = Xrm.Page.getAttribute("ccx_overallragstatus").getValue();
//var serverUrl = Xrm.Page.context.getClientUrl() + '/WebResources/';

//alert ("RAG = " + Rag);

//    if (Rag != 4){
//    Xrm.Page.getControl("ccx_overallragstatus").removeOption(4);
//    }

//    if (Rag == 1) {
//    Xrm.Page.getControl("IFRAME_RAG").setSrc(serverUrl + 'ccx_rag_status_red');
//    }
//
//    if (Rag == 2) {
//    Xrm.Page.getControl("IFRAME_RAG").setSrc(serverUrl + 'ccx_rag_status_amber');
//    }

//    if (Rag == 3) {
//    Xrm.Page.getControl("IFRAME_RAG").setSrc(serverUrl + 'ccx_rag_status_green');
//    }

//    if (Rag == 4) {
//    Xrm.Page.getControl("IFRAME_RAG").setSrc(serverUrl + 'ccx_rag_status_notset');
//    }

//Xrm.Page.getControl("IFRAME_RAG").setVisible(true); // comented out untill live server is patched
//alert ("Color Set");

  // Set Rag Date
    if (Trigger == 'OnChange') {
    Xrm.Page.getAttribute("ccx_overallragstatusdate").setValue(new Date());
    Xrm.Page.getAttribute("ccx_overallragstatusdate").setSubmitMode("always");
    }
}


/* warn of comma in field */
function Comma_In_Service_Type()
{
var comma_problem = false;
var uniqueId = "field_has_comma"; // message identifier
    if (Xrm.Page.getAttribute("ccx_invoice_service_type").getValue() != null)
        if (Xrm.Page.getAttribute("ccx_invoice_service_type").getValue().indexOf(',') != -1)
            comma_problem = true;

    if (comma_problem)
        Xrm.Page.ui.setFormNotification("The Invoice Service Type contains a comma ", "WARNING",uniqueId);
    else
       Xrm.Page.ui.clearFormNotification(uniqueId);
}
