var dhxWins,win,winframe
var opt_net,opt_config

function load_CANsetup()
{
var dhxForm,formStructure
formStructure = [

    {type:"settings",position:"label-top"},
    {type: "fieldset",name:"cansetup", label: "Can Setup", list:[
      {type: "combo", label: "Net", name: "combonet", options:[
      {text: "1", value: "0",selected: true},
      {text: "2", value: "1" }
      ]},
      {type:"newcolumn"},
      {type: "combo", label: "Configuration", name: "comboconfig",  inputLeft:50,  options:[
      {text: "Upstream", value: "0", selected: true},
      {text: "DownStream", value: "1" }
      ]},
      {type:"button", name:"Connect",width:100,offsetTop:10,offsetLeft:100, value:"Connect"}
    ]}
];
dhxWins = new dhtmlXWindows();
//dhxWins.attachViewportTo("Layer_CanSetup");
win = dhxWins.createWindow("cansetup", 100, 100, 500 , 200);
win.setText("CAN Setup");
win.attachURL("CanSetup.html");
win.center();
//win.keepInViewport();
winframe = win.getFrame();
};


dhtmlxEvent(window,"load",function()
{
  LogGrid.load_messagebox();
  load_CANsetup();  
  Layer_TabStrip.style.display = "none";
  Layer_MessageLog.style.display = "none";
});

function onclick_btncanconnect()
{
  opt_config = winframe.contentWindow.document.getElementById("opt_config").value;
  opt_net = winframe.contentWindow.document.getElementById("opt_cannet").value;
  win.close();
}
/*
var formData = [
		{type: "combo", name: "myCombo", label: "Select Band", options:[
				{value: "opt_a", text: "Cradle Of Filth"},
				{value: "opt_b", text: "Children Of Bodom", selected:true}
		]},
		{type: "combo", name: "myCombo2", label: "Select Location", options:[
				{value: "1", text: "Open Air"},
				{value: "2", text: "Private Party"}
		]}
];
 */
