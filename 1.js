
var stone_bl = 0;
var border_on = false;

function addBorder(e){
	if(border_on){
		e.target.classList.add('class_lu');
	}
}

function removeBorder(e){
	e.target.classList.remove('class_lu');
}

function clickBorder(e){
	if(border_on){
		console.log(e);
	}
}

window.addEventListener('mousedown',clickBorder);

window.addEventListener('mouseover',addBorder);

window.addEventListener('mouseout',removeBorder);

window.addEventListener('keydown',function(e){
	//console.log(e.which);
	
	stone_bl += e.which;
	
	console.log(stone_bl);
	
//	if(e.which != 38){  // not start with up, zero it
//		stone_bl = 0;
//	}
	
	if(stone_bl == (38+38+40+40+37+39+37+39+66+65+66+65)){  // yes it is
		border_on = true;
		stone_bl = 0;
	}
	
	if(e.which == 27){ //esc
		border_on = false;
		stone_bl = 0;
	}
	// 38 40 37 39 66 65
})


