var xlsx  = require('xlsx');

module.exports = {

	getTimeTable: function(ws){
		//this function will get the time hashtable from source worksheet

		var timeTable = {};
		var range = xlsx.utils.decode_range(ws['!ref']);

    	for(var C = 5; C <= range.e.c; C++) {
      		var cell_address = {c:C, r:1};
      		var cell = ws[xlsx.utils.encode_cell(cell_address)];
      		if(cell !== undefined && cell.v !== "TOTAL"){
        		timeTable[C] = new Date(cell.w);//.toLocaleDateString();
      		}else if (cell === undefined){
      			continue;
      		}else{
      			break;
      		}
    	}
		return timeTable;
	},

	getTeamTable: function(ws){
		var teamTable = {};
		var range = xlsx.utils.decode_range(ws['!ref']);

		var start;

		for(var R = range.s.r; R<=range.e.r; R++){
			var cell_address = {c:0,r:R};
			var cell = ws[xlsx.utils.encode_cell(cell_address)];
			if(cell !== undefined && cell.v === "Staffing Plan"){
				start = R;
				break;
			}
		}

		for(var R = start+1; R<=range.e.r; R++){
			var cell_address = {c:0,r:R};
			var cell = ws[xlsx.utils.encode_cell(cell_address)];
			if(cell !== undefined && cell.v !== "Onsite"){
				var qualification = ws[xlsx.utils.encode_cell({c:2,r:R})];
				var timezone = ws[xlsx.utils.encode_cell({c:3,r:R})];
				teamTable[R] = [qualification.v,timezone.v];
			}else if(cell === undefined){
				continue;
			}else{
				break;
			}
		}
		return teamTable;
	},

	getOnsiteColor: function(ws){
		var onsite_color = "";
		var range = xlsx.utils.decode_range(ws['!ref']);
		for(var R = range.s.r; R<=range.e.r; R++){
			var cell_address = {c:0,r:R};
			var cell = ws[xlsx.utils.encode_cell(cell_address)];
			if(cell !== undefined && cell.v === "Onsite"){
				onsite_color = cell.s.fgColor.rgb;
				break;
			}
		}
		return onsite_color;
	},

	getRemoteColor: function(ws){
		var onsite_color = "";
		var range = xlsx.utils.decode_range(ws['!ref']);
		for(var R = range.s.r; R<=range.e.r; R++){
			var cell_address = {c:0,r:R};
			var cell = ws[xlsx.utils.encode_cell(cell_address)];
			if(cell !== undefined && cell.v === "Remote"){
				onsite_color = cell.s.fgColor.rgb;
				break;
			}
		}
		return onsite_color;
	},

	prodMem: function(index, onsite_color, remote_color){
	// this function gets onsite/remote member according to colour of cell
		if (index === onsite_color){
			return "Onsite Team Member";
		}
		else if (index === remote_color){
			return "Remote Team Member";
		}
	},

	prodID: function(index, onsite_color, remote_color){
	// this function sets product id number according to colour of cell
		if (index === onsite_color)
		{
			return "9501641";
		}
		else if( index === remote_color)
		{
			return "9501642";
		}	
	},

	teamID: function(team){
	//this function gets a team's id number
		if (team ==="Delivery Team CoE US"){
			return "4183128";
		}
		else if(team === "Delivery Team CoE EMEA Ireland"){
			return "4258806";
		}
		else if( team ==="Delivery Team CoE North Asia"){
			return "12816165";
		}
	},

	itemNum: function(prevQ, curQ, prevT, curT, num){
	// this function determines item line number value
		if ((prevQ === curQ) && (prevT === curT)){
		return (num + 1);
		}
		else return (parseInt(num/10)+1)*10;
	},

	getEndDate: function(date, quantity){
		var numQuantity = parseInt(quantity, 10);
		if(numQuantity === 0) return "TBD";
		var dayOfMonth = date.getDate();
		var endDate = new Date(date.toLocaleDateString());
		endDate.setDate(dayOfMonth + numQuantity -1);
		return (endDate.getMonth()+1+ "/" + endDate.getDate());
	},

	getStartTime: function(team, isGL){
		if (team === "Delivery Team CoE US"){	
			if( isGL === false){
				var time  = new Date("1/1/2015 ,09:00:00");
				return (time.getHours() + ":00");
			}else if (isGL === true){
				var time  = new Date("1/1/2015 ,12:00:00");
				return(time.getHours() + ":00");	
			}
		}else if(team === "Delivery Team CoE EMEA Ireland"){
			var time  = new Date("1/1/2015 ,04:00:00");
			return(time.getHours() + ":00");
		}else if(team ==="Delivery Team CoE North Asia"){
			var time  = new Date("1/1/2015 ,20:00:00");
			return(time.getHours() + ":00");
		}
	},

	getEndTime: function(team, isGL){
		if (team === "Delivery Team CoE US"){
			if(isGL === false ){
				var time  = new Date("1/1/2015 ,18:00:00");
				return(time.getHours() + ":00");
			}else if (isGL === true){
				var time  = new Date("1/1/2015 ,20:00:00");
				return(time.getHours() + ":00");
			}
		} else if(team === "Delivery Team CoE EMEA Ireland"){
			var time  = new Date("1/1/2015 ,12:00:00")
			return(time.getHours() + ":00");
		}else if(team === "Delivery Team CoE North Asia"){
			var time  = new Date("1/1/2015 ,04:00:00")
			return(time.getHours() + ":00");
		}
	},

	isGolive: function(row, col, ws, teamTable){
		
		var curQ = teamTable[row][0];
		var curT = teamTable[row][1];
		
		//case 1
		if(ws[xlsx.utils.encode_cell({c:col,r:row+1})]!==undefined 
			&& ws[xlsx.utils.encode_cell({c:col,r:row+2})] !== undefined
			&& teamTable[row+1] !== undefined 
			&& teamTable[row+2] !== undefined){

			var q1 = teamTable[row+1][0];
			var q2 = teamTable[row+2][0];
			var t1 = teamTable[row+1][1];
			var t2 = teamTable[row+2][1];

			if(curQ === q1 && q1 === q2 && curT !== t1 && t1 !== t2 && curT !== t2){
				return true;
			}

		}

		if(ws[xlsx.utils.encode_cell({c:col,r:row-1})] !== undefined 
			&& ws[xlsx.utils.encode_cell({c:col,r:row+1})] !== undefined
			&& teamTable[row-1] !== undefined 
			&& teamTable[row+1] !== undefined){

			var q1 = teamTable[row-1][0];
			var q2 = teamTable[row+1][0];
			var t1 = teamTable[row-1][1];
			var t2 = teamTable[row+1][1];

			if(curQ === q1 && q1 === q2 && curT !== t1 && t1 !== t2 && curT !== t2){
				return true;
			}

		}

		if(ws[xlsx.utils.encode_cell({c:col,r:row-1})] !== undefined 
			&& ws[xlsx.utils.encode_cell({c:col,r:row-2})] !== undefined
			&& teamTable[row-1] !== undefined
			&& teamTable[row-2] !== undefined){

			var q1 = teamTable[row-1][0];
			var q2 = teamTable[row-2][0];
			var t1 = teamTable[row-1][1];
			var t2 = teamTable[row-2][1];

			if(curQ === q1 && q1 === q2 && curT !== t1 && t1 !== t2 && curT !== t2){
				return true;
			}

		}

		return false;
	}
};