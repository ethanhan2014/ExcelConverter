var express     =    require("express");
var multer      =    require('multer');
var app         =    express();
var fs          =    require('fs');
var xlsx        =    require('xlsx');
var done        =    false;
var newname     =    "";
var reader      =    require('./bo_read.js');
var writer      =    require('./bo_write.js');

app.get('/',function(req,res,next){
      res.sendFile(__dirname+"/index.html");
});


app.post('/upload',[
  multer({ dest: './uploads/'}),

  function(req, res){
  
  try{
  //Read uploaded excel file
  var srcWb = xlsx.readFile(req.files.thefile.path,{'cellStyles':true});
  //Make Time Hashtable and Team Hashtable

  var srcWs = srcWb.Sheets["Engagement Plan"];
  
  //decode excel range
  var range = xlsx.utils.decode_range(srcWs['!ref']);
  
  var timeTable = reader.getTimeTable(srcWs);
  var teamTable = reader.getTeamTable(srcWs);
  var onsite_color = reader.getOnsiteColor(srcWs);
  var remote_color = reader.getRemoteColor(srcWs);

  //Writing into cusomized information

  function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
  }

  var tgwb = new Workbook();

  var startCol = parseInt(Object.keys(timeTable)[0]);

  var startRow = parseInt(Object.keys(teamTable)[0]);
  var endRow = parseInt(Object.keys(teamTable)[Object.keys(teamTable).length-1]);

  for(var num = 1; num<parseInt(Object.keys(timeTable).length/4)+2; num++){ //num<parseInt(Object.keys(timeTable).length/4)
    
    var data = writer.getHeader();
 
    var ws_name = "SO "+num;
    var prevQ = "";
    var curQ  = "";
    var prevT = "";
    var curT  = "";
    var itemNum   = 10;  // set initial item number to be 10

    for(var R = startRow; R <= endRow; R++){
      for(var C = startCol; C < startCol+4; C++) {
        var cell_address = {c:C, r:R};
        var cell = srcWs[xlsx.utils.encode_cell(cell_address)]
        if(cell !== undefined){
          var newrow = [];

          curQ = teamTable[R][0];   //current qualification
          curT = teamTable[R][1];   // current team

          var itemNum = reader.itemNum(prevQ,curQ,prevT,curT,itemNum);
          newrow.push(itemNum); //item number

          newrow.push(20);  //higher number

          newrow.push(cell.v);    // quantity

          newrow.push("TA")       // quantity unit

          newrow.push(cell.v);    // Duration

          newrow.push("DAY");     //unit of duration

          var prodid = reader.prodID(cell.s.fgColor.rgb,onsite_color,remote_color);
          newrow.push(prodid);    //product ID

          var product = reader.prodMem(cell.s.fgColor.rgb,onsite_color,remote_color);
          newrow.push(product);   //product name

          newrow.push(curQ);     // qualification

          var startDate = timeTable[C];
          var sdstr = startDate.getMonth()+1+"/"+startDate.getDate();
          newrow.push(sdstr);         // Requested Start date

          var isgl = reader.isGolive(R,C,srcWs,teamTable);

          var startTime = reader.getStartTime(curT,isgl);
          newrow.push(startTime);         // Request Start Time
        
          newrow.push(reader.getEndDate(startDate,cell.v));         // Requested End Date

          var endTime = reader.getEndTime(curT,isgl);
          newrow.push(endTime);         // Request End Time

          newrow.push("EST");         // Time Zone

          newrow.push(curT);      // service team

          var teamNum = reader.teamID(curT);
          newrow.push(teamNum)

          data.push(newrow);

          prevQ = curQ;
          prevT = curT;
        }
      }
    }


    if(data[2]===undefined) data.push([]);
    data[2][1] = 10; //change higher item number in third row to be 10
    data[2][6] = "";
    data[2][7] = "<SERVICE NAME> Session"

    var tgws = writer.sheet_from_array_of_arrays(data);

    tgwb.SheetNames.push(ws_name);
    tgwb.Sheets[ws_name] = tgws;
    
    startCol = startCol + 4;

  }
  }
  catch(err){
    res.redirect(__dirname+'/error.html');
  }
  xlsx.writeFile(tgwb,'export.xlsx');

  //set the file ready for downloading
  res.setHeader('Content-disposition', 'attachment; filename=export.xlsx');
  res.download('./export.xlsx');

  //Delete the uploaded file in the server
  fs.unlink(req.files.thefile.path,function(err){
    if(err) throw err;
    console.log("Successfully deleted uploaded file");
  });

}]);

app.listen(3000,function(){
    console.log("Working on port 3000");
});