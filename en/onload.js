
  var dhxWins,win,winframe,dhxForm,formStructure,menu
  var opt_net,opt_config

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
  dhxWins.load_cansetup = function()
  {
    win = dhxWins.createWindow("cansetup", 100, 100, 500 , 200);
    win.button("minmax1").hide();
    win.button("park").hide();
    win.button("close").hide();
    win.setText("CAN Setup");
    win.attachURL("CanSetup.html");
    win.center();
    winframe = win.getFrame();
  }
  dhxWins.canconnect = function()
  {
    opt_config = winframe.contentWindow.document.getElementById("opt_config").value;
    opt_net = winframe.contentWindow.document.getElementById("opt_cannet").value;
    win.close();
  }
  var LogGrid = new dhtmlXGridObject('MessageLogObj');
  LogGrid.load_messagebox = function()
  {
    this.setHeader("Date,Time,Information");
    this.setImagePath("../../../codebase/grid/imgs/");
    this.setInitWidths( "100,100,*");
    this.setColAlign ("center,center,left");
    this.setColTypes ("ro,ro,ro");
    this.setColSorting ("na,na,na");
    this.setSkin ("red_gray");
    this.setStyle("color:black;font-weight:bold;","color:black;","","")
    this.enableTooltips ("true,true,true");
    this.enableResizing ("true,true,true");
    this.enableMultiselect(true);
    this.enableAutoWidth (true);
    this.enableSmartRendering(true);
    this.init();
  }


  var SCIGrid = new dhtmlXGridObject('SCILogObj');
  SCIGrid.load_messagebox = function()
  {
    this.setHeader("Date,Time,Information");
    this.setImagePath("../../../codebase/grid/imgs/");
    this.setInitWidths( "100,100,*");
    this.setColAlign ("center,center,left");
    this.setColTypes ("ro,ro,ro");
    this.setColSorting ("na,na,na");
    this.setSkin ("red_gray");
    this.setStyle("color:black;font-weight:bold;","color:black;","","")
    this.enableTooltips ("true,true,true");
    this.enableResizing ("true,true,true");
    this.enableMultiselect(true);
    this.enableAutoWidth (true);
    this.enableSmartRendering(true);
    this.init();
  }
  var MBEEPROMGrid = new dhtmlXGridObject('Layer_MBGrid');
  MBEEPROMGrid.drawgrid = function()
  {
    this.setHeader("Parameter,Address,Data");
    this.setImagePath("../../../codebase/grid/imgs/");
    this.setInitWidths( "200,100,*");
    this.setColAlign ("center,center,left");
    this.setColTypes ("ro,ro,ro");
    this.setColSorting ("na,na,na");
    this.setSkin ("red_gray");
    this.setStyle("color:black;font-weight:bold;","color:black;","","")
    this.enableTooltips ("true,true,true");
    this.enableResizing ("true,true,true");
    this.enableMultiselect(true);
    this.enableAutoWidth (false);
    this.enableSmartRendering(false);
    this.init();
  }
  MBEEPROMGrid.setVal = function(row,val) { this.cells(this.getRowId(row),2).cell.innerHTML = val; }
  MBEEPROMGrid.setCellRed = function(row){this.setCellTextStyle(row,2,"color:red");}
  MBEEPROMGrid.setCellBlack = function(row){this.setCellTextStyle(row,2,"color:Black");}


  var CMEEPROMGrid = new dhtmlXGridObject('Layer_CMGrid');
  CMEEPROMGrid.drawgrid = function()
  {
    this.setHeader("Parameter,Address,Data");
    this.setImagePath("../../../codebase/grid/imgs/");
    this.setInitWidths( "150,100,*");
    this.setColAlign ("center,center,left");
    this.setColTypes ("ro,ro,ro");
    this.setColSorting ("na,na,na");
    this.setSkin ("red_gray");
    this.setStyle("color:black;font-weight:bold;","color:black;","","")
    this.enableTooltips ("true,true,true");
    this.enableResizing ("true,true,true");
    this.enableMultiselect(true);
    this.enableAutoWidth (false);
    this.enableSmartRendering(false);
    this.init();
  }
  CMEEPROMGrid.setVal = function(row,val) { this.cells(this.getRowId(row),2).cell.innerHTML = val; }
  CMEEPROMGrid.setCellRed = function(row){this.setCellTextStyle(row,2,"color:red");}
  CMEEPROMGrid.setCellBlack = function(row){this.setCellTextStyle(row,2,"color:Black");}


  tabbar2 = new dhtmlXTabBar("Layer_TabStripLog","top");
  tabbar2.Init = function()
  {
  tabbar2.setSkin("silver");
  tabbar.setSkinColors("gray","black","#97A0A5");  
  this.setSize("350","500");
  tabbar2.enableAutoReSize( true );
  tabbar2.setImagePath("../../../codebase/tabbar/imgs/");
  tabbar2.addTab("msg_tab1","Message Log");
  tabbar2.addTab("msg_tab2","SCI Log");
  tabbar2.setContent( "msg_tab1", Layer_Tab1_MainLog);
  tabbar2.setContent( "msg_tab2", Layer_Tab2_SCILog);

  for ( var i = 1; i <= tabbar2.getNumberOfTabs(); i++ )
  {
    tabbar2.setCustomStyle( 'msg_tab' + i, 'gray', 'black', 'font-size:10pt;font-family:Arial;font-weight: bold;' );
  }
  tabbar2.setTabActive("msg_tab1");
  }
  tabbar = new dhtmlXTabBar("Layer_TabStripMain","top");
  tabbar.Init = function() 
  {
  tabbar.enableAutoReSize( true );
  tabbar.tabstyle
  tabbar.setImagePath("../../../codebase/tabbar/imgs/");
  tabbar.setSkin("silver");  
  tabbar.setSkinColors("gray","black","#97A0A5");
  tabbar.addTab("main_tab1","Main");
  tabbar.addTab("main_tab2","Status");
  tabbar.addTab("main_tab3","CM EEPROM");
  tabbar.addTab("main_tab4","MB EEPROM");
  tabbar.setContent( "main_tab1", Layer_Tab1_Main);
  tabbar.setContent( "main_tab2", Layer_Tab2_Status);
  tabbar.setContent( "main_tab3", Layer_Tab3_CMEEPROM);
  tabbar.setContent( "main_tab4", Layer_Tab4_MBEEPROM);

  for ( var i = 1; i <= tabbar.getNumberOfTabs(); i++ )
  {
    tabbar.setCustomStyle( 'main_tab' + i, 'gray', 'black', 'font-size:10pt;font-family:Arial;font-weight: bold;' );
  }
  
  tabbar.setTabActive("main_tab3");
  }