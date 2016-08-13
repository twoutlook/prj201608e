/**
 * Enterprise Spreadsheet Solutions
 * Copyright(c) FeyaSoft Inc. All right reserved.
 * info@enterpriseSheet.com
 * http://www.enterpriseSheet.com
 * 
 * Licensed under the EnterpriseSheet Commercial License.
 * http://enterprisesheet.com/license.jsp
 * 
 * You need to have a valid license key to access this file.
 */
Ext.onReady(function () {

    Ext.QuickTips.init();
    /**
     * Define those 2 methods as global variable
     */
    SHEET_API = Ext.create('EnterpriseSheet.api.SheetAPI', {
        openFileByOnlyLoadDataFlag: true
    });

    //DOING LAYOUT
    SHEET_API_HD = SHEET_API.createSheetApp({
//        withoutTitlebar: false,
        withoutTitlebar: true,
        withoutSheetbar: false,
//        withoutToolbar: false,
        withoutToolbar: true,
        withoutContentbar: false,
//        withoutContentbar: true,
//        withoutSidebar: false
        withoutSidebar: true

    });
    // this is tab panel include main and details 
    var centralPanel = Ext.create('enterpriseSheet.templates.CenterPanel', {
    });
    Ext.create('Ext.Viewport', {
        layout: 'border',
        items: [centralPanel],
        listeners: {
            afterlayout: function (v, layout, eOpts) {
                // westPanel.selectNode();
            }
        }
    });
    //  NOTE:
    //  
    var colorStep = "#A9BCF5";
    var colorStepEnd = "#E6E6E6";
    var colorSect = "#837E7C"; //bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true
    var colorDdl = "#F9E79F"; //#82E0AA  
    var colorInput = "#F4D03F"; // 
    var colorVersion = "#98AFC7"; // 

//
    //
    //
    //
//        var ddlMaterial = {data: "xxx", drop: Ext.encode({data: "xxx,yyy,zzz"})};
    var ddlMaterial = {bgc: colorDdl, ta: "center", data: "===材质规格===", drop: Ext.encode({data: "一。铝合金,AlSi10Mg,AlSi9Cu3,ADC12,A380,A413,A360,A356,二。锌合金 ,ZINC-2,ZINC-3,ZINC-5,ZINC-7,ZAMAK-8,三。镁合金 ,AZ91D,AM60,THX-AZ91D,THX-AM60"})};


    //挂镀铬，挂镀镍，挂镀沙丁镍，挂镀沙丁铬，
//拉丝镀铬，拉丝镀镍，
//挂镀黑铬，挂镀化学镍，
//挂镀枪黑，镀PVD
//滚镀镍
//滚镀化学镍，挂镀锌，滚镀锌
//    var ddlSurface = {bgc: colorDdl, ta: "center", data: "===表面要求===", drop: Ext.encode({data: "单清洗,烤漆前皮膜（含清洗）,铝合金一般皮膜（48H）,锌合金一般皮膜（48H),镁合金一般皮膜（24H),铝合金-特殊皮膜（   H）,锌合金-特殊皮膜（   H) ,特殊导电皮膜(  欧姆) ,粉體烤漆-A+級,粉體烤漆-A級 ,粉體烤漆-B級 ,液體烤漆-A+級 ,液體烤漆-A級 ,液體烤漆-B級 ,阳极氧化-A级 , 阳极氧化-B级 ,电泳-A级,电泳-B级 ,掛鍍-A級,掛鍍-B 級,滾鍍-A級,滾鍍-B級 ,高清洁度清洗（600um）,高清洁度清洗（400um）,高清洁度清洗（200um）,清洗鉻酸"})};

    var ddlSurface_data = "单清洗,烤漆前皮膜（含清洗）,铝合金一般皮膜（48H）,锌合金一般皮膜（48H),镁合金一般皮膜（24H),铝合金-特殊皮膜（   H）,锌合金-特殊皮膜（   H) ,特殊导电皮膜(  欧姆) ,粉體烤漆-A+級,粉體烤漆-A級 ,粉體烤漆-B級 ,液體烤漆-A+級 ,液體烤漆-A級 ,液體烤漆-B級 ,阳极氧化-A级 , 阳极氧化-B级 ,电泳-A级,电泳-B级 ,掛鍍-A級,掛鍍-B 級,滾鍍-A級,滾鍍-B級 ,高清洁度清洗（600um）,高清洁度清洗（400um）,高清洁度清洗（200um）,清洗鉻酸";



    var A0602_ddlSurface_data = ",挂镀铬,挂镀镍,挂镀沙丁镍,挂镀沙丁铬,拉丝镀铬,拉丝镀镍,挂镀黑铬,挂镀化学镍,挂镀枪黑,镀PVD,滚镀镍,滚镀化学镍,挂镀锌,滚镀锌";

    var ddlSurface = {bgc: colorDdl, ta: "center", data: "===表面要求===", drop: Ext.encode({data: ddlSurface_data + A0602_ddlSurface_data})};

    var ddlMachine = {bgc: colorDdl, ta: "center", data: "===适用机型===", drop: Ext.encode({data: "一。铝合金压铸合模费：,铝-125T,鋁-150T/160T/200T,鋁-250T/280T,鋁-350T/340T/400T,鋁-550T/530T/500T,鋁-630T/650T,鋁-800T/900T,鋁-1250T,鋁-1600T,鋁-2000T,鋁-2500T,鋁-3000T,二。锌合金压铸合模费：,鋅-快速机/4轴机,鋅-10T/5T,鋅-15T/20T,鋅-25T/30T/40T,鋅-50T/60T,鋅-80T/100T/125,鋅-150T,鋅-185/200T,鋅-250T,鋅-300T,三。镁合金压铸合模费：,鎂-150T,鎂-340T/400T,鎂-530T/660T,THX-鎂280T,THX-鎂650T"})};
//DOING...喷沙sand blasting 抛丸.震动研磨
    var ddlSand = {bgc: colorDdl, ta: "center", data: "===喷砂,抛丸,震动研磨===", drop: Ext.encode({data: "喷砂,抛丸,震动研磨"})};
    var ddlCold = {bgc: colorDdl, ta: "center", data: "===冷喷.热烧===", drop: Ext.encode({data: "冷冻去除毛刺,热能去除毛刺"})};
    var ddlStep9 = {bgc: colorDdl, ta: "center", data: "===皮膜处理===", drop: Ext.encode({data: "烤漆前清洗皮膜,一般清洗加皮膜,特殊要求皮膜,清洗,特殊清洗"})};
    var ddlMaching = {bgc: colorDdl, ta: "center", data: "===机加工===", drop: Ext.encode({data: "CNC机加工,传统机加工"})};
    var ddl079 = {bgc: colorDdl, ta: "center", data: "===皮膜处理===", drop: Ext.encode({data: "清洗,一般皮膜(48H),鎂合金一般皮膜(24H),高清洁度清洗(600um),高清洁度清洗(400um),高清洁度清洗(200um),特殊皮膜"})};

    //[[A0602]]
//    var ddl086Data = "粉體烤漆-A+級,粉體烤漆-A級 ,粉體烤漆-B級 ,液體烤漆-A+級 ,液體烤漆-A級 ,液體烤漆-B級 ,阳极氧化-A级 , 阳极氧化-B级 ,电泳-A级,电泳-B级 ,掛鍍-A級,掛鍍-B 級,滾鍍-A級,滾鍍-B級 ,高清洁度清洗（600um）,高清洁度清洗（400um）,高清洁度清洗（200um）,清洗鉻酸";

    //[[A0603]]
    var ddl086Data = "粉體烤漆-A+級,粉體烤漆-A級 ,粉體烤漆-B級 ,液體烤漆-A+級 ,液體烤漆-A級 ,液體烤漆-B級 ,阳极氧化-A级 , 阳极氧化-B级 ,电泳-A级,电泳-B级 ,掛鍍-A級,掛鍍-B 級,滾鍍-A級,滾鍍-B級";
    var ddl086 = {bgc: colorDdl, ta: "center", data: "===表面要求(2)===", drop: Ext.encode({data: ddl086Data})};

    //[[A0606]] 其它特殊处理”一栏需细化，除了会议中提及的渗补、时效
    var ddlSpecialData = "　,渗补,时效";//全形空白
    var ddlSpecial = {bgc: colorDdl, ta: "center", data: "===其它特殊处理===", drop: Ext.encode({data: ddlSpecialData})};

//var colorStep="#E6E6E6";
//var colorStepEnd="#A9BCF5";


    function mergeJSON(source1, source2) {
        /*
         * Properties from the Souce1 object will be copied to Source2 Object.
         * Note: This method will return a new merged object, Source1 and Source2 original values will not be replaced.
         * */
        var mergedJSON = Object.create(source2); // Copying Source2 to a new Object

        for (var attrname in source1) {
            if (mergedJSON.hasOwnProperty(attrname)) {
                if (source1[attrname] != null && source1[attrname].constructor == Object) {
                    /*
                     * Recursive call if the property is an object,
                     * Iterate the object and set all properties of the inner object.
                     */
                    mergedJSON[attrname] = mergeJSON(source1[attrname], mergedJSON[attrname]);
                }

            } else {//else copy the property from source1
                mergedJSON[attrname] = source1[attrname];
            }
        }

        return mergedJSON;
    }

    function styleSubTotal(source2) {
        var subtotal = {bgc: colorStepEnd, fm: "money|¥|2|none", dsd: "ed", cal: true};
        return    mergeJSON(subtotal, source2);
    }
    function styleSubTotalMoney(money, source2) {
        var subtotal = {bgc: colorStepEnd, fm: "money|" + money + "|2|none", dsd: "ed", cal: true};
        return    mergeJSON(subtotal, source2);
    }
    function styleInput(source2) {
        var basic = {bgc: colorInput, ta: "right"};
        return    mergeJSON(basic, source2);
    }

    function setNaToZero(str) {
        return   "=IF(ISNA(" + str + "),0,(" + str + "))";
    }
    function setNaToEmpty(str) {
        return   "=IF(ISNA(" + str + "),,(" + str + "))";
    }



    var json = {
        fileName: "Project201608",
        sheets: [{id: 1, name: "欠料", actived: true, color: "orange"},
        ],
//        floatings: [
//            {sheet: 1, name: "colGroups", ftype: "colgroup", json: "[{level:1, span:[2,3]}, {level:1, span:[4,6]}]"},
//        ],
        cells: [
            {sheet: 1, row: 0, col: 0, json: {fz: 10, ww: 'break-word', ws: 'normal', dsd: "", height: 20, va: "middle"}},
            {sheet: 1, row: 0, col: 1, json: {ta: "center", dsd: "", data: "A", width: 60}},
            {sheet: 1, row: 0, col: 2, json: {ta: "center", data: "B", width: 160}},
            {sheet: 1, row: 0, col: 3, json: {ta: "center",dsd: "", width: 120}},
            {sheet: 1, row: 0, col: 4, json: {ta: "center",dsd: "", width: 120}},
            {sheet: 1, row: 0, col: 5, json: {ta: "center",dsd: "", width: 120}},
            {sheet: 1, row: 0, col: 6, json: {ta: "center",dsd: "", width: 120}},
            {sheet: 1, row: 0, col: 7, json: {ta: "center",dsd: "", width: 120}},
            {sheet: 1, row: 0, col: 8, json: {ta: "center",dsd: "", width: 120}},
          
            // {sheet: 1, row: 3, col: 2, json: {data: "接受询价日期："}},
            // {sheet: 1, row: 3, col: 1, json: {data: "生成", it: "button", btnStyle: "color: blue; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK___MAKE_EXCEL"}},
            // {sheet: 1, row: 4, col: 2, json: {data: "业务担当："}},

// 客戶	品名	欠料量	欠備庫量	客訴量	模具壽命?%	客戶	品名	欠料量	欠備庫量	客訴量	模具壽命?%
            // {sheet: 1, row: 3, col: 1, json: {data: "生成", it: "button", btnStyle: "color: blue; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK___MAKE_EXCEL"}},
            {sheet: 1, row: 3, col: 1, json: {data: "生成", it: "button", btnStyle: "color: blue; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK___MAKE_EXCEL"}},
          
            {sheet: 1, row: 1, col: 2, json: {data: "客戶"}},
            {sheet: 1, row: 1, col: 3, json: {data: "品名"}},
            {sheet: 1, row: 1, col: 4, json: {data: "欠料量"}},
            {sheet: 1, row: 1, col: 5, json: {data: "欠備庫量"}},
            {sheet: 1, row: 1, col: 6, json: {data: "客訴量"}},
            {sheet: 1, row: 1, col: 7, json: {data: "模具壽命?%"}},
            {sheet: 1, row: 2, col: 2, json: {data: "CISCO"}},
            {sheet: 1, row: 3, col: 2, json: {data: "DELTA"}},
     



//
//          // 只是做成最後一行
            // {sheet: 1, row: 7, col: 1, json: {bgc: colorVersion, data: 'RFQ'}},
            // {sheet: 1, row: 8, col: 1, json: {bgc: colorVersion, data: 'A0707'}}, //VERSION
            // {sheet: 1, row: 9, col: 1, json: {bgc: colorVersion, data: '公版'}},
        ]
    };
    // =============================================================================================
    // ok inject data now ...
//    var json_default = {
//		fileName: "Employee Directory",
//		sheets:[{id:1, name:"Main view", actived:true, color:"orange"}],
//		floatings: [
//	        { sheet:1, name:"colGroups", ftype:"colgroup", json: "[{level:1, span:[2,3]}, {level:1, span:[4,6]}]" },
//	    ],
//		cells:[
//		    { sheet: 1, row: 0, col: 0, json: { height: 20, va: "middle"} },
//		    { sheet: 1, row: 0, col: 1, json: { data: "ID", width: 50, dcfg: "{dt:0, io:true, min:0, max:10000, op:0, ignoreBlank: true, titleIcon: \"number\"}", ticon:"number" } },
//			{ sheet: 1, row: 0, col: 2, json: { data: "Name", width: 100, ticon:"profile"} },
//			{ sheet: 1, row: 0, col: 3, json: { data: "Dept.(Remote)", width: 130, drop: "list", dcfg: "{dt:15, url: \"fakeData/dropdownList\", titleIcon:  \"remoteList\"}", ticon:"remoteList" } },
//			{ sheet: 1, row: 0, col: 4, json: { data: "Email", width: 110, dcfg: "{dt:9}", ticon:"email" } },
//			{ sheet: 1, row: 0, col: 5, json: { data: "Phone", width: 100, dcfg: "{dt:8}", ticon:"phone" } },
//			{ sheet: 1, row: 0, col: 6, json: { data: "Gender", width: 80, drop: "list", dcfg: "{dt:13, list: [\"Male\",\"Female\"]}", ticon:"dropdown" } },
//			{ sheet: 1, row: 0, col: 7, json: { data: "Birth date", width: 120, drop: "date", fm: "date", dfm: "F d, Y", ticon:"calendar"  } },			
//			{ sheet: 1, row: 0, col: 8, json: { data: "Contact picker", width: 170, ticon:"contact", beforeEdit: "_beforeeditcell_" } },
//			{ sheet: 1, row: 0, col: 9, json: { data: "Manager?", width: 100, it: "checkbox", itchk: false, ta: "center", ticon:"checkbox" } },
//			{ sheet: 1, row: 0, col: 10, json: { data: "Images", width: 130, dcfg: "{dt:7}", ticon:"image" } },
//			{ sheet: 1, row: 0, col: 11, json: { data: "Salary", dcfg: "{dt:11, format: \"money|$|2|none|usd|true\"}",  ticon:"money_dollar" } },
//			{ sheet: 1, row: 0, col: 12, json: { data: "Percent", dcfg: "{dt:12, format: \"0.00%\"}",  ticon:"percent" } },
//			{ sheet: 1, row: 0, col: 13, json: { data: "Notes", dcfg: "{dt:14, titleIcon: \"textLong\"}",  ticon:"textLong" } },
//			
//			{ sheet: 1, row: 1, col: 1, json: { data: 1 } },
//			{ sheet: 1, row: 1, col: 2, json: { data: 'Jerry Marc' } },
//			{ sheet: 1, row: 1, col: 3, json: { render:'dropRender', data: 'HR Dept', dropId: 1} },
//			{ sheet: 1, row: 1, col: 4, json: { data: 'john.marc@abc.com'} },
//			{ sheet: 1, row: 1, col: 5, json: { data: '1 (888) 456-7654'} },
//			{ sheet: 1, row: 1, col: 6, json: { data: 'Female'} },
//			{ sheet: 1, row: 1, col: 7, json: { data: '1982-01-15', fm: "date", dfm: "F d, Y" } },
//			{ sheet: 1, row: 1, col: 8, json: { render:'contactRender', data: "Eva Mat, John Marc", itms: '[{name: "Eva Mat", email: "eva@gmail.com", id: 8}, {name: "John Marc", email: "john@abc.com", id: 9}]' } },
//			{ sheet: 1, row: 1, col: 10, json: { render:'attachRender', itms: '[{aid: "rT7KfpHA8cI_", url: "sheetAttach/downloadFile?attachId=rT7KfpHA8cI_", type: "img", name: "blue.jpg"},{aid: "2ZisVQ1-*Lo_", url: "sheetAttach/downloadFile?attachId=2ZisVQ1-*Lo_", type: "img", name: "green.jpg"}]' } },
//			{ sheet: 1, row: 1, col: 11, json: { data: 82334.5678 } },
//			{ sheet: 1, row: 1, col: 12, json: { data: 0.96 } },
//			{ sheet: 1, row: 1, col: 13, json: { data: 'This is notes, it is a long text. Double click to edit it.' } },
//			
//			{ sheet: 1, row: 2, col: 1, json: { data: 2 } },
//			{ sheet: 1, row: 2, col: 2, json: { data: 'Dave Smith' } },
//			{ sheet: 1, row: 2, col: 3, json: { render:'dropRender', data: 'Software Dept', dropId: 2} },
//			{ sheet: 1, row: 2, col: 4, json: { data: 'dave.smith@abc.com'} },
//			{ sheet: 1, row: 2, col: 5, json: { data: '1 (888) 231-7654'} },
//			{ sheet: 1, row: 2, col: 6, json: { data: 'Male'} },
//			{ sheet: 1, row: 2, col: 7, json: { data: '1980-01-15', fm: "date", dfm: "F d, Y" } },
//			{ sheet: 1, row: 2, col: 8, json: { render:'contactRender', data: "Christina Angela, Marina Chris", itms: '[{name: "Christina Angela", email: "christina@gmail.com", id: 4}, {name: "Marina Chris", email: "marina@abc.com", id: 6}]' } },
//			{ sheet: 1, row: 2, col: 10, json: { render:'attachRender', itms: '[{aid: "CIBHu3ffG8Q_", url: "sheetAttach/downloadFile?attachId=CIBHu3ffG8Q_", type: "img", name: "admin.png"},{aid: "VcrhEYAyrzA_", url: "sheetAttach/downloadFile?attachId=VcrhEYAyrzA_", type: "img", name: "asset.png"}]' } },
//			{ sheet: 1, row: 2, col: 11, json: { data: 81234.5678 } },
//			{ sheet: 1, row: 2, col: 12, json: { data: 0.95 } },
//			{ sheet: 1, row: 2, col: 13, json: { data: 'This is notes, it is a long text. Double click to edit it.' } },
//			
//			{ sheet: 1, row: 3, col: 1, json: { data: 3 } },
//			{ sheet: 1, row: 3, col: 2, json: { data: 'Kevin Featherstone' } },
//			{ sheet: 1, row: 3, col: 3, json: { render:'dropRender', data: 'Software Dept', dropId: 2} },
//			{ sheet: 1, row: 3, col: 4, json: { data: 'kevin@abc.com'} },
//			{ sheet: 1, row: 3, col: 5, json: { data: '1 (888) 232-7654'} },
//			{ sheet: 1, row: 3, col: 6, json: { data: 'Male'} },
//			{ sheet: 1, row: 3, col: 7, json: { data: '1990-01-15', fm: "date", dfm: "F d, Y" } },
//			{ sheet: 1, row: 3, col: 8, json: { render:'contactRender', data: "Christina Angela, Marina Chris", itms: '[{name: "Christina Angela", email: "christina@gmail.com", id: 4}, {name: "Marina Chris", email: "marina@abc.com", id: 6}]' } },
//			{ sheet: 1, row: 3, col: 10, json: { render:'attachRender', itms: '[{aid: "CIBHu3ffG8Q_", url: "sheetAttach/downloadFile?attachId=CIBHu3ffG8Q_", type: "img", name: "admin.png"},{aid: "VcrhEYAyrzA_", url: "sheetAttach/downloadFile?attachId=VcrhEYAyrzA_", type: "img", name: "asset.png"}]' } },
//			{ sheet: 1, row: 3, col: 11, json: { data: 81934.5678 } },
//			{ sheet: 1, row: 3, col: 12, json: { data: 0.98 } },
//			{ sheet: 1, row: 3, col: 13, json: { data: 'This is notes, it is a long text. Double click to edit it.' } }
//		]
//	};	



    // [[A0601]]
    // NOTE: use system function to insert row
    //SHEET_API.loadData(SHEET_API_HD, json, null, this);
    var MAX_ROW_NUMBER = 120;
    var MAX_COL_NUMBER = 9;

    SHEET_API.loadData(SHEET_API_HD, json, function () {
        SHEET_API.setMaxRowNumber(MAX_ROW_NUMBER);
//        SHEET_API.setMaxColNumber(8);//col H is NOT working well when copy/paste col
        SHEET_API.setMaxColNumber(MAX_COL_NUMBER);

    }, this);

//http://www.enterprisesheet.com/api/docs/manageDataAPIs/sheetAPI.html#insertRow
//insertRow( hd, sheetId, row, rowSpan )


//    SHEET_API.insertRow(SHEET_API_HD, 1, 11,1);
//    SHEET_API.insertRow(SHEET_API_HD, 1, 11,1);

    /**
     * [[A0601]]-insert row and assign data (end)
     * 
     */


    /*
     * 1.10-5)插入 遮蔽費用
     2.10-6)插入 印刷費用
     3.10-7)插入 镭雕費用
     4.10)步驟和小計加入"印刷镭雕"字樣,調整小計公式
     5.最後加 合計(含稅),公式 X1.17
     6.最後加 運費,預設值空白
     7.模具寿命, 預設值空白
     8.年需求量下一行, 加MOQ,預設值空白
     9.加工夾治具費用後加, 專用測試設備費用,同時調整模具總價公式
     10.建兵將提供︰表面要求,電鍍部份細化
     11.文清將整理︰其它特殊處理項目和費率
     */

    // [[A0601]] 
    //   5.最後加 合計(含稅),公式 X1.17
    //   6.最後加 運費,預設值空白
    SHEET_API.updateCells(SHEET_API_HD, getPatchCellA001(1));
    console.log("PATCH#1 after SHEET_API.updateCells(SHEET_API_HD,  getPatchCellA001(1));---合計(含稅), 運費");

    /**
     * [[A0601]]
     1.10-5)插入 遮蔽費用
     2.10-6)插入 印刷費用
     3.10-7)插入 镭雕費用
     4.10)步驟和小計加入"印刷镭雕"字樣,調整小計公式
     */
    SHEET_API.insertRow(SHEET_API_HD, 1, 89, 1);
    SHEET_API.insertRow(SHEET_API_HD, 1, 89, 1);
    SHEET_API.insertRow(SHEET_API_HD, 1, 89, 1);
    SHEET_API.updateCells(SHEET_API_HD, getPatchCellA001(2));
    console.log("PATCH#2 after SHEET_API.updateCells(SHEET_API_HD,  getPatchCellA001(2));---遮蔽,印刷,镭雕");


    SHEET_API.insertRow(SHEET_API_HD, 1, 22, 1);
    SHEET_API.updateCells(SHEET_API_HD, getPatchCellA001(3));
    console.log("PATCH#3 after SHEET_API.updateCells(SHEET_API_HD,  getPatchCellA001(3));---专用测试设备费用");
//    console.log(getPatchCellA001(3));

//                                      getPatchCellA0602



    function getTemp1() {

//        var ddl086Data = "粉體烤漆-A+級,粉體烤漆-A級 ,粉體烤漆-B級 ,液體烤漆-A+級 ,液體烤漆-A級 ,液體烤漆-B級 ,阳极氧化-A级 , 阳极氧化-B级 ,电泳-A级,电泳-B级 ,掛鍍-A級,掛鍍-B 級,滾鍍-A級,滾鍍-B級 ,高清洁度清洗（600um）,高清洁度清洗（400um）,高清洁度清洗（200um）,清洗鉻酸";
//        var ddl086 = {bgc: colorDdl, ta: "center", data: "===表面要求(2)===", drop: Ext.encode({data: ddl086Data})};

        var cells = [];
        cells.push(
                {sheet: 1, row: 85, col: 2, json: {bgc: colorStep, data: "10)粉体烤漆.液体烤漆...： "}},
                {sheet: 1, row: 86, col: 3, json: ddl086}, // 
                {sheet: 1, row: 86, col: 4, json: ddl086}, // 
                {sheet: 1, row: 86, col: 5, json: ddl086}, // 
                {sheet: 1, row: 86, col: 6, json: ddl086}, // 
                {sheet: 1, row: 86, col: 7, json: ddl086}, // 
                {sheet: 1, row: 86, col: 8, json: ddl086}, // 
        //
        //
        // LAST ONE ============================================================================================
                {sheet: 1, row: 1, col: 1, json: {data: ""}}
        );
        return cells;
    }

    console.log("PATCH#4 ---");
    SHEET_API.updateCells(SHEET_API_HD, getTemp1());
    console.log("PATCH#4 after SHEET_API.updateCells(SHEET_API_HD,  getPatchCellA001(4));--- 086 單獨下拉");
    // console.log(getPatchCellsA0602_1());

    console.log("PATCH A0602");
    SHEET_API.updateCells(SHEET_API_HD, getPatchA0602(1));

    //[[A0603]]
    console.log("PATCH A0603");
    SHEET_API.updateCells(SHEET_API_HD, getPatchA0603(1));

    //WARNING NOT WORKING 查表取2,3 ok, 4不行!!!???
    //
    //[[A0603]] ROW14, MOQ, LOOKUP
    SHEET_API.updateCells(SHEET_API_HD, getPatchA0603(2));



    function getPatchA0606InMain() {
        var cells = [];
        cells.push(
                {sheet: 1, row: 96, col: 3, json: ddlSpecial}, // 
                {sheet: 1, row: 96, col: 4, json: ddlSpecial}, // 
                {sheet: 1, row: 96, col: 5, json: ddlSpecial}, // 
                {sheet: 1, row: 96, col: 6, json: ddlSpecial}, // 
                {sheet: 1, row: 96, col: 7, json: ddlSpecial}, // 
                {sheet: 1, row: 96, col: 8, json: ddlSpecial} // 
        );
        return cells;
    }
    //[[A0606]]  吴文清 根据2016.5.30号讨论“其它特殊处理”一栏需细化，除了会议中提及的渗补、时效外，一时想不起来，还有哪些？
    SHEET_API.updateCells(SHEET_API_HD, getPatchA0606InMain());
    //
    //文清發現 #95 公式錯誤
    //SHEET_API.updateCells(SHEET_API_HD, getPatchA0606(1));
    //

//    SHEET_API.loadData(SHEET_API_HD, json, function () {
//        SHEET_API.deleteRow(SHEET_API_HD, 1, 5,0); //[[A0606]] CLEAN UP UNUSED ROW, BY 文清
//    }, this);


    SHEET_API.setFocus(SHEET_API_HD, 3, 1);



//
//

// add event listener - this shows the code to add customer function 
    var sheet = SHEET_API_HD.sheet;
    var editor = sheet.getEditor();
    editor.on('quit', function (editor, sheetId, row, col) {
        if (col === 1) {
            // this is the method to query customer existing backend and auto fill data
            // 
            // NOTE, GOOD EXAMPLE, MAYBE IMPLEMENT IT LATER 05/06 15:32
//            var employeeId = SHEET_API.getCellValue(SHEET_API_HD, sheetId, row, col).data;
//            if (employeeId)
//                AUTO_FILL_CUSTOMER_DATA_BY_EMPLOYEEID(employeeId, sheetId, row, col);
        }
    }, this);
    // add cell on select event ...
    /**
     var sm = sheet.getSelectionModel();
     sm.on('selectionchange', function(startPos, endPos, region, sm) {
     if (startPos.row == endPos.row && startPos.col == endPos.col && startPos.col == 8) {
     this.customEditor = Ext.create('customer.CellEditor', {
     sheetId: region.sheetId,
     row: startPos.row,
     col: startPos.col
     });
     this.customEditor.popup();
     }
     }, this);
     **/

    sheet.on('_beforeeditcell_', function (sheetId, row, col, cellData, sheet) {
        var contactWin = Ext.create('customer.CellEditor', {
            sheetId: sheetId,
            row: row,
            col: col,
            cellData: cellData,
            sheet: sheet
        });
        contactWin.popup();
        return false;
    }, this);
});
	