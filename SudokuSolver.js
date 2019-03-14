var FSO = WScript.CreateObject("Scripting.FileSystemObject");

var RES_NOT_ACCOMPLISHED = 0;
var RES_SUCCESS = 1;
var RES_FAILURE = 2;

var ST_NORMAL = 1;
var ST_SAMURAI = 2;
var ST_CLIPBOARD = 3;
var ST_DESIGN = 4;

var ppBorderBottom = 3;
var ppBorderLeft = 2;
var ppBorderRight = 4;
var ppBorderTop = 1;

var RGB_GREEN = 0x50b000;

String.prototype.getNumbers = function() {
	var result = "";
	for (var i = 0;i < this.length;i++)
		if (!isNaN(parseInt(this.charAt(i))))
			result += this.charAt(i);
	return(result);
};

String.prototype.getFirstNumber = function() {
	for (var i = 0;i < this.length;i++)
		if (!isNaN(parseInt(this.charAt(i))))
			return this.charAt(i);
	
	return "";
};

checkScript();

WScript.StdOut.WriteLine("Welcome to Sudoku Solver.");

var sudokuType;
checkArguments();

//Obtain an HTML document object.
var document = WScript.CreateObject("htmlfile");

if (sudokuType == ST_NORMAL)
{
	//Create our sudoku table.
	var sudoku = document.appendChild(document.createElement("body")).appendChild(document.createElement("table"));
	//Add nine rows and nine columns for
	//our table.
	for (var i = 0;i < 9;i++)
	{
		var tr = sudoku.insertRow();
		for (var j = 0;j < 9;j++)
			tr.insertCell();
	}

	//Take a sudoku from the user.
	var sudokuNumbers = getSudoku();
	//If the user has entered invalid data,
	//display an error.
	if (!sudokuNumbers)
	{
		WScript.StdErr.WriteLine("Invalid sudoku.");
		askToClose();
	}
	//Initialize our sudoku table with
	//the given data. Also, ignore the
	//zero numbers that mean the 
	//corresponding cell is blank.
	for (i = 0;i < sudokuNumbers.length;i++)
		if (parseInt(sudokuNumbers.charAt(i)))
			sudoku.cells[i].innerText = sudokuNumbers.charAt(i);
	//////////////////
	markBlankCells();
	//////////////////
	//Solve the sudoku
	tryToSolve();
	//Let the user press Enter to 
	//close the program.
	askToClose(sudokuType);
}
else if (sudokuType == ST_SAMURAI)
{
	var $body = document.appendChild(document.createElement("body"));
	
	var $centralGrid = $createGrid();
	var $upperLeftGrid = $createGrid();
	var $upperRightGrid = $createGrid();
	var $lowerLeftGrid = $createGrid();
	var $lowerRightGrid = $createGrid();
	
	$associateRelativeSquares();
	
	$getSudoku();
	
	markBlankCells();
	
	$tryToSolve();
	
	askToClose(sudokuType);
}
else if (sudokuType == ST_CLIPBOARD)
{
	WScript.StdOut.WriteLine("Not currently supported.");
	askToClose();
}
else if (sudokuType == ST_DESIGN)
{
	WScript.StdOut.WriteLine("Setting up a new presentation in PowerPoint...");
	
	var PP = createPP();
	
	var prs = PP.Presentations.Add(0);
	
	var sld1 = prs.Slides.Add(1, 11);
	sld1.Shapes.Title.TextFrame.TextRange.Text = "Central Grid";
	
	var tblCentralGrid = createSudokuTable(sld1);
	setTextFormatting(tblCentralGrid);
	
	var sld2 = prs.Slides.Add(2, 11);
	sld2.Shapes.Title.TextFrame.TextRange.Text = "Upper-left Grid";
	
	var tblUpperLeftGrid = createSudokuTable(sld2);
	setTextFormatting(tblUpperLeftGrid);
	
	ulg_removeBorders(tblUpperLeftGrid.Cell(7, 7));
	ulg_removeBorders(tblUpperLeftGrid.Cell(7, 8));
	ulg_removeBorders(tblUpperLeftGrid.Cell(7, 9));
	ulg_removeBorders(tblUpperLeftGrid.Cell(8, 7));
	ulg_removeBorders(tblUpperLeftGrid.Cell(8, 8));
	ulg_removeBorders(tblUpperLeftGrid.Cell(8, 9));
	ulg_removeBorders(tblUpperLeftGrid.Cell(9, 7));
	ulg_removeBorders(tblUpperLeftGrid.Cell(9, 8));
	ulg_removeBorders(tblUpperLeftGrid.Cell(9, 9));
	
	var sld3 = prs.Slides.Add(3, 11);
	sld3.Shapes.Title.TextFrame.TextRange.Text = "Upper-right Grid";
	
	var tblUpperRightGrid = createSudokuTable(sld3);
	setTextFormatting(tblUpperRightGrid);
	
	urg_removeBorders(tblUpperRightGrid.Cell(7, 1));
	urg_removeBorders(tblUpperRightGrid.Cell(8, 1));
	urg_removeBorders(tblUpperRightGrid.Cell(9, 1));
	urg_removeBorders(tblUpperRightGrid.Cell(7, 2));
	urg_removeBorders(tblUpperRightGrid.Cell(8, 2));
	urg_removeBorders(tblUpperRightGrid.Cell(9, 2));
	urg_removeBorders(tblUpperRightGrid.Cell(7, 3));
	urg_removeBorders(tblUpperRightGrid.Cell(8, 3));
	urg_removeBorders(tblUpperRightGrid.Cell(9, 3));
	
	var sld4 = prs.Slides.Add(4, 11);
	sld4.Shapes.Title.TextFrame.TextRange.Text = "Lower-left Grid";
	
	var tblLowerLeftGrid = createSudokuTable(sld4);
	setTextFormatting(tblLowerLeftGrid);
	
	llg_removeBorders(tblLowerLeftGrid.Cell(1, 7));
	llg_removeBorders(tblLowerLeftGrid.Cell(2, 7));
	llg_removeBorders(tblLowerLeftGrid.Cell(3, 7));
	llg_removeBorders(tblLowerLeftGrid.Cell(1, 8));
	llg_removeBorders(tblLowerLeftGrid.Cell(2, 8));
	llg_removeBorders(tblLowerLeftGrid.Cell(3, 8));
	llg_removeBorders(tblLowerLeftGrid.Cell(1, 9));
	llg_removeBorders(tblLowerLeftGrid.Cell(2, 9));
	llg_removeBorders(tblLowerLeftGrid.Cell(3, 9));
	
	var sld5 = prs.Slides.Add(5, 11);
	sld5.Shapes.Title.TextFrame.TextRange.Text = "Lower-right Grid";
	
	var tblLowerRightGrid = createSudokuTable(sld5);
	setTextFormatting(tblLowerRightGrid);
	
	lrg_removeBorders(tblLowerRightGrid.Cell(1, 1));
	lrg_removeBorders(tblLowerRightGrid.Cell(1, 2));
	lrg_removeBorders(tblLowerRightGrid.Cell(1, 3));
	lrg_removeBorders(tblLowerRightGrid.Cell(2, 1));
	lrg_removeBorders(tblLowerRightGrid.Cell(2, 2));
	lrg_removeBorders(tblLowerRightGrid.Cell(2, 3));
	lrg_removeBorders(tblLowerRightGrid.Cell(3, 1));
	lrg_removeBorders(tblLowerRightGrid.Cell(3, 2));
	lrg_removeBorders(tblLowerRightGrid.Cell(3, 3));
	
	PP.Activate();
	prs.NewWindow().WindowState = 3;
	
	WScript.StdOut.WriteLine("Now you can enter your sudoku in the");
	WScript.StdOut.WriteLine("PowerPoint window created by this app.");
	WScript.StdOut.WriteLine("If you have a normal 9x9 sudoku, then ");
	WScript.StdOut.WriteLine("enter it in the first slide and delete ");
	WScript.StdOut.WriteLine("the other slides.");
	WScript.StdOut.WriteLine("If you have a Samurai sudoku, then enter ");
	WScript.StdOut.WriteLine("each grid in each corresponding slide.");
	WScript.StdOut.WriteLine("When you are finished with PowerPoint, ");
	WScript.StdOut.WriteLine("just press Enter right in the console window.");
	WScript.StdIn.SkipLine();
	
	WScript.StdOut.WriteLine("Processing...");
	switch (prs.Slides.Count)
	{
	case 1 :
		sudokuType = ST_NORMAL;
		
		var sudoku = document.appendChild(document.createElement("body")).appendChild(document.createElement("table"));
		convertSudoku();
		
		prs.Close();
		PP.Quit();
		
		markBlankCells();
		
		tryToSolve();
		
		askToClose(sudokuType);
		break;
	case 5 :
		sudokuType = ST_SAMURAI;
		
		var $body = document.appendChild(document.createElement("body"));
		
		var $centralGrid = $createGrid();
		var $upperLeftGrid = $createGrid();
		var $upperRightGrid = $createGrid();
		var $lowerLeftGrid = $createGrid();
		var $lowerRightGrid = $createGrid();
		
		$associateRelativeSquares();
		
		$convertCentralGrid();
		$convertGrid(tblUpperLeftGrid, $upperLeftGrid);
		$convertGrid(tblUpperRightGrid, $upperRightGrid);
		$convertGrid(tblLowerLeftGrid, $lowerLeftGrid);
		$convertGrid(tblLowerRightGrid, $lowerRightGrid);
		
		prs.Close();
		PP.Quit();
		
		markBlankCells();
		
		$tryToSolve();
		
		askToClose(sudokuType);
		
		break;
	}
}

//Functions for normal 9x9 sudokus
function tryToSolve()
{
	WScript.StdOut.WriteLine("Solving... Please wait...");
	
	if (!_tryToSolve())
	{
		WScript.StdErr.WriteLine("This sudoku is incorrect or has errors!!!!!");
		WScript.StdErr.WriteLine("So it cannot be solved!!");
		return;
	}
	
	if (isFull())	//If the sudoku is fully solved,
	{
		//Display the solved sudoku.
		WriteOutputInfo();
		//Exit function.
		return;
	}

	WScript.StdOut.WriteLine("Still trying... Please wait...");
	
	var d = findDoubt();
	if (!d.invokeGuesses())
	{
		WScript.StdErr.WriteLine("This sudoku is incorrect or has errors!!!!!");
		WScript.StdErr.WriteLine("So it cannot be solved!!");
		return;
	}
	
	//Write the result in Standard Output
	WriteOutputInfo();
}

function _tryToSolve()
{
	try
	{
		//Look in the columns, rows, and squares for 
		//numbers, until the sudoku becomes full, or
		//the program becomes unable to find any more 
		//numbers.
		do
			if (isFull())
				break;
		while (lookInSquares() || lookInRows() || lookInColumns())
		
		return true;
	}
	catch (e)
	{
		return false;
	}
}

function lookInSquares()
{
	//The following function checks whether or
	//not a number can be placed in a cell. It
	//takes a cell and a number as parameters,
	//and scans the cell's parent row or column,
	//in order to check for possible conflicts
	//(a number in a cell that is the same as 
	//the number passed as parameter)
	function checkForConflicts(cell, number)
	{
		//Scan the cell's parent row.
		var cells = cell.parentElement.cells;
		for (var i = 0;i < cells.length;i++)
			if (cells[i].innerText == number)
				return false;
		//Scan the cell's parent column.
		for (var i = 0;i < sudoku.rows.length;i++)
			if (sudoku.rows[i].cells[cell.cellIndex].innerText == number)
				return false;
		//If it comes here, there are no
		//conflicts, so return true.
		return true;
	}
	
	var foundOne = false;	//Variable containing a boolean value
							//indicating whether at least one 
							//number is discovered.
	//Loop that examines the squares one by one.
	for (var i = 0;i < 9;i++)
	{
		//Call getSquare to obtain an array of
		//cells in the corresponding square.
		var square = getSquare(i);
		//Call filterFullCells to obtain an 
		//array of blank cells.
		var blankCells = filterFullCells(square);
		
		loopNumbers:	//Loop that examines the square and
						//determines which number does not 
						//exist in the square and then places
						//that number in an appropriate cell.
		for (var j = 1;j <= 9;j++)
		{
			//Check if j exists in the square.
			//In that case, go to the next iteration.
			for (var k in square)
				if (square[k].innerText == j)
					continue loopNumbers;

			var possibleCells = []; //Array that contains the cells
									//that the number j can be possibly
									//placed in.
			for (k in blankCells)
				if (checkForConflicts(blankCells[k], j))
					possibleCells.push(blankCells[k]);

			//If there's more than one possible cell,
			//go to the next iteration.
			if (possibleCells.length > 1)
				continue;
			
			//Place the number j in the cell.
			possibleCells[0].innerText = j;
			
			//Because we discovered a number, set
			//foundOne variable to true.
			foundOne=true;
		}
	}
	
	return(foundOne);
}

function lookInRows()
{
	var foundOne = false;	//Variable containing a boolean value
							//indicating whether at least one
							//number is discovered.
	//Loop that examines rows one by one
	for (var i = 0;i < 9;i++)
	{
		//Obtain an array of cells in the
		//corresponding row.
		var row = sudoku.rows[i].cells;
		//Call filterFullCells to obtain
		//an array of blank cells.
		var blankCells = filterFullCells(row);
		
		loopNumbers:	//Loop that examines the row and
						//determines which number does not
						//exist in the row and then places 
						//that number in an appropriate cell.
		for (var j = 1;j <= 9;j++)
		{
			//Check if j exists in the row.
			//In that case, go to the next
			//iteration.
			for (var k = 0;k < 9;k++)
				if (row[k].innerText == j)
					continue loopNumbers;

			var possibleCells = [];	//Array that contains the cells
									//that the number j can be possibly
									//placed in.
			for (k in blankCells)
				if (checkForConflicts(blankCells[k], j))
					possibleCells.push(blankCells[k]);

			//If there is more than one possible cell,
			//go to the next iteration.
			if (possibleCells.length > 1)
				continue;
			
			//Place the number j in the cell.
			possibleCells[0].innerText = j;
			
			//Because we discovered a number,
			//set foundOne variable to true.
			foundOne = true;
		}
	}
	
	return(foundOne);
	
	//The following function checks whether or
	//not a number can be placed in a cell. It
	//takes a cell and a number as parameters,
	//and scans the cell's parent square or 
	//column in order to check for possible
	//conflicts (a number in a cell that is the
	//same as the number passed as parameter)
	function checkForConflicts(cell, number)
	{
		//Find the cell's parent square.
		var square = whichSquare(cell);
		//Scan the cell's parent square.
		for (var i in square)
			if (square[i].innerText == number)
				return false;
		//Scan the cell's parent column.
		for (i=0;i<9;i++)
			if (sudoku.rows[i].cells[cell.cellIndex].innerText == number)
				return false;
		//If it comes here, there are no
		//conflicts, so return true.
		return true;
	}
}

function lookInColumns()
{
	//By seeing the comments within the lookInRows
	//function, you will find out how this function
	//works, as well.
	
	var foundOne = false;
	
	for (var i = 0;i < 9;i++)
	{
		loopNumbers:
		for (var j = 1;j <= 9;j++)
		{
			for (var k = 0;k < 9;k++)
				if (sudoku.rows[k].cells[i].innerText == j)
					continue loopNumbers;

			var possibleCells = [];
			for (k = 0;k < 9;k++)
				if ((!sudoku.rows[k].cells[i].innerText) && checkForConflicts(sudoku.rows[k].cells[i], j))
					possibleCells.push(sudoku.rows[k].cells[i]);

			if (possibleCells.length > 1)
				continue;
			
			possibleCells[0].innerText = j;
			
			foundOne = true;
		}
	}
	return foundOne;

	function checkForConflicts(cell, number)
	{
		var square = whichSquare(cell);
		for (var i in square)
			if (square[i].innerText == number)
				return false;
		
		var row = cell.parentElement.cells;
		for (i = 0;i < 9;i++)
			if (row[i].innerText == number)
				return false;
		
		return true;
	}
}

function getSquare(index)
{
	//Increment index to turn it into a 1-based index.
	index++;
	//Calculate the coordinates of the square.
	var top = Math.ceil(index / 3);
	var left = index + 3 - 3 * top;
	//Calculate the coordinates of the first cell in 
	//the square.
	var cl = 3 * left - 2;
	var ct = 3 * top - 2;
	//Calculate the zero-based index of the first cell 
	//in the square.
	var ci = 9 * ct + cl - 10;
	//Define an array to store the indexes of the cells 
	//in the square.
	var square = [ci, ci + 1, ci + 2];
	square.push(square[0] + 9, square[1] + 9, square[2] + 9,
				square[0] + 18, square[1] + 18, square[2] + 18);
	//Store the real cells in our array instead of 
	//indexes
	for (var i in square)
		square[i] = sudoku.cells[square[i]];
	//Return our array
	return square;
}

function whichSquare(cell)
{
	//Calculate the coordinates of the cell's parent square.
	var top = Math.ceil((cell.parentElement.rowIndex + 1) / 3);
	var left = Math.ceil((cell.cellIndex + 1) / 3);
	//Calculate the zero-based index of the cell's parent
	//square, then pass the index to getSquare function to 
	//obtain an array of cells in the square, and finally
	//return that array.
	return getSquare(3 * top + left - 4);
}

function getSudoku()
{
	WScript.StdOut.WriteLine("Please enter your sudoku.");
	WScript.StdOut.WriteLine("---------");
	//Let the user enter some numbers representing
	//ether the content of first row of the sudoku 
	//or the content of the entire sudoku.
	var sudokuNumbers = WScript.StdIn.ReadLine();
	//If the user did not enter anything, just quit
	//the app.
	if (!sudokuNumbers)
		WScript.Quit();
	//Check the length of the input string to see
	//if it equals 81, meaning that the contents 
	//of the entire sudoku is entered. Then return
	//the whole string.
	if (sudokuNumbers.length == 81)
		return sudokuNumbers;
	//Otherwise if the length of the string equals
	//nine, it means that the contents of only one
	//row is entered. So we should take the content
	//of other rows as well.
	else if (sudokuNumbers.length == 9)
	{
		for (var i = 0;i < 8;i++)
		{
			//Let the user enter some numbers representing
			//the contents of the row.
			var line = WScript.StdIn.ReadLine().getNumbers();
			//If the user has not entered nine numbers, we
			//have some invalid data so we should return 
			//null.
			if (line.length != 9)
				return null;
			//Store the numbers in our variable
			sudokuNumbers += line;
		}
		return sudokuNumbers;
	}
	//If the length of the input string is neither
	//nine nor eighty-one, the user has entered 
	//invalid data so we should return null.
	else
		return null;
}

function WriteOutputInfo(grid)
{
	if (!grid)
		grid = sudoku;
	
	with (WScript.StdOut)
	{
		Write("-------------");
		
		for (var i = 0;i < grid.rows.length;i++)
		{
			if (Column !== 1)
				Write("\n");
			
			Write("|");
			
			for (var j = 0;j < grid.rows[i].cells.length;j++)
			{
				var n = grid.rows[i].cells[j].innerText;
				Write(isNaN(parseInt(n)) ? "0" : n);
				
				if (!((j + 1) % 3))
					Write("|");
			}
			
			if (!((i + 1) % 3))
				WriteLine("\n-------------");
		}
	}
}

function findDoubt()
{
	function checkForConflicts(cell, number)
	{
		var cells = cell.parentElement.cells;
		for (var i = 0;i < cells.length;i++)
			if (cells[i].innerText == number)
				return false;
		
		for (var i = 0;i < sudoku.rows.length;i++)
			if (sudoku.rows[i].cells[cell.cellIndex].innerText == number)
				return false;

		return true;
	}
	
	for (var i = 0;i < 9;i++)
	{
		var square = getSquare(i);
		var blankCells = filterFullCells(square);
		loopNumbers:
		for (var j = 1;j <= 9;j++)
		{
			for (var k in square)
				if (square[k].innerText == j)
					continue loopNumbers;

			var possibleCells = [];
			for (k in blankCells)
				if (checkForConflicts(blankCells[k], j))
					possibleCells.push(blankCells[k]);
			
			return new Doubt(j, possibleCells);
		}
	}
}

//Functions for Samurai Sudokus

function $createGrid()
{
	var grid = $body.appendChild(document.createElement("table"));
	for (var i = 0;i < 9;i++)
	{
		var tr = grid.insertRow();
		for (var j = 0;j < 9;j++)
			tr.insertCell();
	}
	return grid;
}

function $getSquare(grid, index)
{
	var top = Math.ceil((++index) / 3);
	var left = index + 3 - 3 * top;
	
	var cl = 3 * left - 2;
	var ct = 3 * top - 2;

	var ci = 9 * ct + cl - 10;

	var square = [ci, ci + 1, ci + 2];
	square.push(square[0] + 9, square[1] + 9, square[2] + 9,
				square[0] + 18, square[1] + 18, square[2] + 18);

	for (var i in square)
		square[i] = grid.cells[square[i]];

	return square;
}

function $whichSquare(cell)
{
	var top = Math.ceil((cell.parentElement.rowIndex + 1) / 3);
	var left = Math.ceil((cell.cellIndex + 1) / 3);
	return $getSquare(cell.parentElement.parentElement.parentElement, 3 * top + left - 4);
}

function $associateRelativeSquares()
{
	$associateRelativeCells($getSquare($upperLeftGrid, 8), $getSquare($centralGrid, 0));
	$associateRelativeCells($getSquare($upperRightGrid, 6), $getSquare($centralGrid, 2));
	$associateRelativeCells($getSquare($lowerLeftGrid, 2), $getSquare($centralGrid, 6));
	$associateRelativeCells($getSquare($lowerRightGrid, 0), $getSquare($centralGrid, 8));
}

function $associateRelativeCells(sq1, sq2)
{
	for (var i in sq1)
	{
		sq1[i].relative = sq2[i];
		sq2[i].relative = sq1[i];
	}
}

function $getSudoku()
{
	function get_separate(l)
	{
		for (var i = 0;i < l.length;i++)
			if (parseInt(l.charAt(i)))
			{
				var cell = $centralGrid.rows[0].cells[i];
				cell.innerText = l.charAt(i);
				if (cell.relative)
					cell.relative.innerText = l.charAt(i);
			}

		for (var i = 1;i < 9;i++)
		{
			var line = WScript.StdIn.ReadLine().getNumbers();
			
			if (line.length != 9)
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			for (var j = 0;j < 9;j++)
				if (parseInt(line.charAt(j)))
				{
					var cell = $centralGrid.rows[i].cells[j];
					cell.innerText = line.charAt(j);
					if (cell.relative)
						cell.relative.innerText = line.charAt(j);
				}
		}
		
		WScript.StdOut.WriteLine("The upper-left grid");
		for (var i = 0;i < 9;i++)
		{
			var line = WScript.StdIn.ReadLine().getNumbers();
			
			if (i > 5)
			{
				if ((line.length != 9) && (line.length != 6))
				{
					WScript.StdErr.WriteLine("Invalid sudoku.");
					askToClose();
				}
				
				line = line.substr(0, 6);
			}
			else
				if (line.length != 9)
				{
					WScript.StdErr.WriteLine("Invalid sudoku.");
					askToClose();
				}
			
			for (var j = 0;j < line.length;j++)
				if (parseInt(line.charAt(j)))
					$upperLeftGrid.rows[i].cells[j].innerText = line.charAt(j);
		}
		
		WScript.StdOut.WriteLine("The upper-right grid");
		for (var i = 0;i < 9;i++)
		{
			if (i > 5)
			{
				var iRow = i - 6;
				WScript.StdOut.Write(cellToNumber($centralGrid.rows[iRow].cells[6]));
				WScript.StdOut.Write(cellToNumber($centralGrid.rows[iRow].cells[7]));
				WScript.StdOut.Write(cellToNumber($centralGrid.rows[iRow].cells[8]));
			}
			var line = WScript.StdIn.ReadLine().getNumbers();
			
			if (line.length != (i > 5 ? 6 : 9))
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			for (var j = 0;j < line.length;j++)
				if (parseInt(line.charAt(j)))
					$upperRightGrid.rows[i].cells[(i > 5) ? (j + 3) : j].innerText = line.charAt(j);
		}
		
		WScript.StdOut.WriteLine("The lower-left grid");
		for (var i = 0;i < 9;i++)
		{
			var line = WScript.StdIn.ReadLine().getNumbers();
			
			if (i < 3)
			{
				if ((line.length != 9) && (line.length != 6))
				{
					WScript.StdErr.WriteLine("Invalid sudoku.");
					askToClose();
				}
				
				line = line.substr(0, 6);
			}
			else
				if (line.length != 9)
				{
					WScript.StdErr.WriteLine("Invalid sudoku.");
					askToClose();
				}
			
			for (var j = 0;j < line.length;j++)
				if (parseInt(line.charAt(j)))
					$lowerLeftGrid.rows[i].cells[j].innerText = line.charAt(j);
		}
		
		WScript.StdOut.WriteLine("The lower-right grid");
		for (var i = 0;i < 9;i++)
		{
			if (i < 3)
			{
				var iRow = i + 6;
				WScript.StdOut.Write(cellToNumber($centralGrid.rows[iRow].cells[6]));
				WScript.StdOut.Write(cellToNumber($centralGrid.rows[iRow].cells[7]));
				WScript.StdOut.Write(cellToNumber($centralGrid.rows[iRow].cells[8]));
			}
			
			var line = WScript.StdIn.ReadLine().getNumbers();
			
			if (line.length != (i < 3 ? 6 : 9))
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			for (var j = 0;j < line.length;j++)
				if (parseInt(line.charAt(j)))
					$lowerRightGrid.rows[i].cells[(i < 3) ? (j + 3) : j].innerText = line.charAt(j);
		}
	}
	
	function get_standard(l)
	{
		for (var i = 0;i < 9;i++)
			if (parseInt(l.charAt(i)))
				$upperLeftGrid.rows[0].cells[i].innerText = l.charAt(i);
		
		for (i = 9;i < 18;i++)
			if (parseInt(l.charAt(i)))
				$upperRightGrid.rows[0].cells[i - 9].innerText = l.charAt(i);
		
		for (i = 1;i <= 5;i++)
		{
			var line = WScript.StdIn.ReadLine().getNumbers();
			if (line.length != 18)
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			var str1 = line.substr(0, 9);
			var str2 = line.substr(9);
			
			for (var j = 0;j < 9;j++)
			{
				if (parseInt(str1.charAt(j)))
					$upperLeftGrid.rows[i].cells[j].innerText = str1.charAt(j);
				
				if (parseInt(str2.charAt(j)))
					$upperRightGrid.rows[i].cells[j].innerText = str2.charAt(j);
			}
		}
		
		WScript.StdOut.WriteLine("---------------------");
		for (var i = 0;i < 3;i++)
		{
			var line = WScript.StdIn.ReadLine().getNumbers();
			if (line.length != 21)
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			var str1 = line.substr(0, 9);
			var str2 = line.substr(12);
			var str3 = line.substr(9, 3);
			
			var iRow = i + 6;
			for (var j = 0;j < 9;j++)
			{
				if (parseInt(str1.charAt(j)))
				{
					var cell = $upperLeftGrid.rows[iRow].cells[j];
					cell.innerText = str1.charAt(j);
					if (cell.relative)
						cell.relative.innerText = str1.charAt(j);
				}
				if (parseInt(str2.charAt(j)))
				{
					var cell = $upperRightGrid.rows[iRow].cells[j];
					cell.innerText = str2.charAt(j);
					if (cell.relative)
						cell.relative.innerText = str2.charAt(j);
				}
			}
			
			for (j = 0;j < 3;j++)
				if (parseInt(str3.charAt(j)))
					$centralGrid.rows[i].cells[j + 3].innerText = str3.charAt(j);
		}
		
		WScript.StdOut.WriteLine("      ---------      ");
		for (i = 3;i < 6;i++)
		{
			WScript.StdOut.Write("      ");
			var line = WScript.StdIn.ReadLine().getNumbers();
			if (line.length != 9)
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			for (var j = 0;j < 9;j++)
				if (parseInt(line.charAt(j)))
					$centralGrid.rows[i].cells[j].innerText = line.charAt(j);
		}
		
		WScript.StdOut.WriteLine("---------------------");
		for (i = 0;i < 3;i++)
		{
			var line = WScript.StdIn.ReadLine().getNumbers();
			if (line.length != 21)
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			var str1 = line.substr(0, 9);
			var str2 = line.substr(12);
			var str3 = line.substr(9, 3);
			
			for (j = 0;j < 9;j++)
			{
				if (parseInt(str1.charAt(j)))
				{
					var cell = $lowerLeftGrid.rows[i].cells[j];
					cell.innerText = str1.charAt(j);
					if (cell.relative)
						cell.relative.innerText = str1.charAt(j);
				}
				if (parseInt(str2.charAt(j)))
				{
					var cell = $lowerRightGrid.rows[i].cells[j];
					cell.innerText = str2.charAt(j);
					if (cell.relative)
						cell.relative.innerText = str2.charAt(j);
				}
			}
			
			for (var j = 0;j < 3;j++)
				if (parseInt(str3.charAt(j)))
					$centralGrid.rows[i + 6].cells[j + 3].innerText = str3.charAt(j);
		}
		
		WScript.StdOut.WriteLine("---------   ---------");
		for (i = 3;i < 9;i++)
		{
			var line = WScript.StdIn.ReadLine().getNumbers();
			if (line.length != 18)
			{
				WScript.StdErr.WriteLine("Invalid sudoku.");
				askToClose();
			}
			
			var str1 = line.substr(0, 9);
			var str2 = line.substr(9);
			
			for (j = 0;j < 9;j++)
			{
				if (parseInt(str1.charAt(j)))
					$lowerLeftGrid.rows[i].cells[j].innerText = str1.charAt(j);
				
				if (parseInt(str2.charAt(j)))
					$lowerRightGrid.rows[i].cells[j].innerText = str2.charAt(j);
			}
		}
	}
	
	WScript.StdOut.WriteLine("The central grid or the entire sudoku.");
	WScript.StdOut.WriteLine("---------   ---------");
	var line = WScript.StdIn.ReadLine().getNumbers();
	switch (line.length)
	{
	case 9 :
		get_separate(line);
		break;
	case 18 :
		get_standard(line);
		break;
	default :
		WScript.StdErr.WriteLine("Invalid sudoku.");
		askToClose();
	}
}

function $lookInSquares(grid)
{
	function checkForConflicts(cell, number)
	{
		var cells = cell.parentElement.cells;
		for (var i = 0;i < cells.length;i++)
			if (cells[i].innerText == number)
				return false;

		for (i = 0;i < grid.rows.length;i++)
			if (grid.rows[i].cells[cell.cellIndex].innerText == number)
				return false;

		if (cell.relative)
		{
			cells = cell.relative.parentElement.cells;
			for (i = 0;i < cells.length;i++)
				if (cells[i].innerText == number)
					return false;
			
			var relativeGrid = cell.relative.parentElement.parentElement.parentElement;
			for (i = 0;i < relativeGrid.rows.length;i++)
				if (relativeGrid.rows[i].cells[cell.relative.cellIndex].innerText == number)
					return false;
		}
		
		return true;
	}
	
	var foundOne = false;
	for (var i = 0;i < 9;i++)
	{
		var square = $getSquare(grid, i);
		var blankCells = filterFullCells(square);
		
		loopNumbers:
		for (var j = 1;j <= 9;j++)
		{
			for (var k in square)
				if (square[k].innerText == j)
					continue loopNumbers;
				
			var possibleCells = [];
			for (k in blankCells)
				if (checkForConflicts(blankCells[k], j))
					possibleCells.push(blankCells[k]);
				
			if (possibleCells.length > 1)
				continue;
			
			possibleCells[0].innerText = j;
			if (possibleCells[0].relative)
				possibleCells[0].relative.innerText = j;
			
			foundOne = true;
		}
	}
	
	return foundOne;
}

function $lookInRows(grid)
{
	function checkForConflicts(cell, number)
	{
		var square = $whichSquare(cell);
		for (var i in square)
			if (square[i].innerText == number)
				return false;
			
		for (i=0;i<9;i++)
			if (grid.rows[i].cells[cell.cellIndex].innerText == number)
				return false;
			
		if (cell.relative)
		{
			var relativeRow = cell.relative.parentElement.cells;
			for (i = 0;i < relativeRow.length;i++)
				if (relativeRow[i].innerText == number)
					return false;
			
			var relativeGrid = cell.relative.parentElement.parentElement.parentElement;
			for (i = 0;i < relativeGrid.rows.length;i++)
				if (relativeGrid.rows[i].cells[cell.relative.cellIndex].innerText == number)
					return false;
		}
		
		return true;
	}
	
	var foundOne = false;
	
	for (var i = 0;i < 9;i++)
	{
		var row = grid.rows[i].cells;
		var blankCells = filterFullCells(row);
		
		loopNumbers:
		for (var j = 1;j <= 9;j++)
		{
			for (var k = 0;k < 9;k++)
				if (row[k].innerText == j)
					continue loopNumbers;
			
			var possibleCells = [];
			for (k in blankCells)
				if (checkForConflicts(blankCells[k], j))
					possibleCells.push(blankCells[k]);
			
			if (possibleCells.length > 1)
				continue;
			
			possibleCells[0].innerText = j;
			if (possibleCells[0].relative)
				possibleCells[0].relative.innerText = j;
			
			foundOne = true;
		}
	}
	
	return foundOne;
}

function $lookInColumns(grid)
{
	function checkForConflicts(cell, number)
	{
		var square = $whichSquare(cell);
		for (var i in square)
			if (square[i].innerText == number)
				return false;
		
		var row = cell.parentElement.cells;
		for (i = 0;i < row.length;i++)
			if (row[i].innerText == number)
				return false;
		
		if (cell.relative)
		{
			var relativeRow = cell.relative.parentElement.cells;
			for (i = 0;i < relativeRow.length;i++)
				if (relativeRow[i].innerText == number)
					return false;
			
			var relativeGrid = cell.relative.parentElement.parentElement.parentElement;
			for (i = 0;i < relativeGrid.rows.length;i++)
				if (relativeGrid.rows[i].cells[cell.relative.cellIndex].innerText == number)
					return false;
		}
		
		return true;
	}
	
	var foundOne = false;
	
	for (var i = 0;i < 9;i++)
	{
		loopNumbers:
		for (var j = 1;j <= 9;j++)
		{
			for (var k = 0;k < 9;k++)
				if (grid.rows[k].cells[i].innerText == j)
					continue loopNumbers;
				
			var possibleCells = [];
			for (k = 0;k < 9;k++)
				if ((!grid.rows[k].cells[i].innerText) && checkForConflicts(grid.rows[k].cells[i], j))
					possibleCells.push(grid.rows[k].cells[i]);
				
			if (possibleCells.length > 1)
				continue;
			
			possibleCells[0].innerText = j;
			if (possibleCells[0].relative)
				possibleCells[0].relative.innerText = j;
			
			foundOne = true;
		}
	}
	
	return foundOne;
}

function $findDoubt()
{
	function lookInSquares(grid)
	{
		function checkForConflicts(cell, number)
		{
			var cells = cell.parentElement.cells;
			for (var i = 0;i < cells.length;i++)
				if (cells[i].innerText == number)
					return false;

			for (i = 0;i < grid.rows.length;i++)
				if (grid.rows[i].cells[cell.cellIndex].innerText == number)
					return false;

			if (cell.relative)
			{
				cells = cell.relative.parentElement.cells;
				for (i = 0;i < cells.length;i++)
					if (cells[i].innerText == number)
						return false;
				
				var relativeGrid = cell.relative.parentElement.parentElement.parentElement;
				for (i = 0;i < relativeGrid.rows.length;i++)
					if (relativeGrid.rows[i].cells[cell.relative.cellIndex].innerText == number)
						return false;
			}
			
			return true;
		}
		
		for (var i = 0;i < 9;i++)
		{
			var square = $getSquare(grid, i);
			var blankCells = filterFullCells(square);
			
			loopNumbers:
			for (var j = 1;j <= 9;j++)
			{
				for (var k in square)
					if (square[k].innerText == j)
						continue loopNumbers;
					
				var possibleCells = [];
				for (k in blankCells)
					if (checkForConflicts(blankCells[k], j))
						possibleCells.push(blankCells[k]);
					
				return new Doubt(j, possibleCells);
			}
		}
		
		return null;
	}
	
	var result = lookInSquares($upperLeftGrid);
	if (!result)
		result = lookInSquares($upperRightGrid);
	if (!result)
		result = lookInSquares($lowerLeftGrid);
	if (!result)
		result = lookInSquares($lowerRightGrid);
	if (!result)
		result = lookInSquares($centralGrid);
	
	return result;
}

function $writeOutputInfo()
{
	function displayRow(cir)
	{
		var arr = [];
		for (var i = 0;i < cir.length;i += 3)
			arr.push(cellToNumber(cir[i]) + cellToNumber(cir[i + 1]) + cellToNumber(cir[i + 2]));
		WScript.StdOut.Write(arr.join("|"));
	}
	
	function combineRows(firstRow, middleRow, lastRow)
	{
		var result = [];
		
		for (var i = 0;i < firstRow.cells.length;i++)
			result.push(firstRow.cells[i]);
		
		result.push(middleRow.cells[3], middleRow.cells[4], middleRow.cells[5]);
		
		for (i = 0;i < lastRow.cells.length;i++)
			result.push(lastRow.cells[i]);
		
		return result;
	}
	
	for (var i = 0;i < 6;i++)
	{
		displayRow($upperLeftGrid.rows[i].cells);
		WScript.StdOut.Write("     ");
		displayRow($upperRightGrid.rows[i].cells);
		WScript.StdOut.Write("\n");
		
		if (!((i + 1) % 3))
			WScript.StdOut.WriteLine("-----------     -----------");
	}
	
	displayRow(combineRows($upperLeftGrid.rows[6], $centralGrid.rows[0], $upperRightGrid.rows[6]));
	WScript.StdOut.Write("\n");
	displayRow(combineRows($upperLeftGrid.rows[7], $centralGrid.rows[1], $upperRightGrid.rows[7]));
	WScript.StdOut.Write("\n");
	displayRow(combineRows($upperLeftGrid.rows[8], $centralGrid.rows[2], $upperRightGrid.rows[8]));
	WScript.StdOut.Write("\n        -----------        \n");
	
	WScript.StdOut.Write("        ");
	displayRow($centralGrid.rows[3].cells);
	WScript.StdOut.Write("        \n");
	WScript.StdOut.Write("        ");
	displayRow($centralGrid.rows[4].cells);
	WScript.StdOut.Write("        \n");
	WScript.StdOut.Write("        ");
	displayRow($centralGrid.rows[5].cells);
	WScript.StdOut.Write("        \n");
	WScript.StdOut.WriteLine("        -----------        ");
	
	displayRow(combineRows($lowerLeftGrid.rows[0], $centralGrid.rows[6], $lowerRightGrid.rows[0]));
	WScript.StdOut.Write("\n");
	displayRow(combineRows($lowerLeftGrid.rows[1], $centralGrid.rows[7], $lowerRightGrid.rows[1]));
	WScript.StdOut.Write("\n");
	displayRow(combineRows($lowerLeftGrid.rows[2], $centralGrid.rows[8], $lowerRightGrid.rows[1]));
	WScript.StdOut.Write("\n-----------     -----------\n");
	
	for (var i = 3;i < 9;i++)
	{
		displayRow($lowerLeftGrid.rows[i].cells);
		WScript.StdOut.Write("     ");
		displayRow($lowerRightGrid.rows[i].cells);
		WScript.StdOut.Write("\n");
		
		if (i == 5)
			WScript.StdOut.WriteLine("-----------     -----------");
	}
}

function $_tryToSolve()
{
	try
	{
		do
			if (isFull())
				break;
		while (	$lookInSquares($centralGrid) ||
				$lookInSquares($upperLeftGrid) ||
				$lookInSquares($upperRightGrid) ||
				$lookInSquares($lowerLeftGrid) ||
				$lookInSquares($lowerRightGrid) ||
				
				$lookInRows($centralGrid) ||
				$lookInRows($upperLeftGrid) ||
				$lookInRows($upperRightGrid) ||
				$lookInRows($lowerLeftGrid) ||
				$lookInRows($lowerLeftGrid) ||
				
				$lookInColumns($centralGrid) ||
				$lookInColumns($upperLeftGrid) ||
				$lookInColumns($upperRightGrid) ||
				$lookInColumns($lowerLeftGrid) ||
				$lookInColumns($lowerLeftGrid))

		return true;
	}
	catch (e)
	{
		return false;
	}
}

function $tryToSolve()
{
	WScript.StdOut.WriteLine("Solving... Please wait...");
	
	if (!$_tryToSolve())
	{
		WScript.StdErr.WriteLine("This sudoku is incorrect or has errors!!!!!");
		WScript.StdErr.WriteLine("So it cannot be solved!!");
		return;
	}
	
	if (isFull())
	{
		$writeOutputInfo();
		return;
	}

	WScript.StdOut.WriteLine("Still trying... Please wait...");
	
	var d = $findDoubt();
	if (!d.invokeGuesses())
	{
		WScript.StdErr.WriteLine("This sudoku is incorrect or has errors!!!!!");
		WScript.StdErr.WriteLine("So it cannot be solved!!");
		return;
	}
	
	$writeOutputInfo();
}

function cellToNumber(cell)
{
	var t = cell.innerText;
	return(t ? t : "0");
}

//PowerPoint Manipulation Functions

function createPP()
{
	try
	{
		return WScript.CreateObject("PowerPoint.Application");
	}
	catch (e)
	{
		WScript.StdErr.WriteLine("Microsoft Office PowerPoint is not installed on your computer.");
		askToClose();
	}
}

function createSudokuTable(slide)
{
	var table = slide.Shapes.AddTable(9, 9, -1, -1, 340, 340).Table;
	
	table.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}");
	table.TableDirection = 1;
	
	table.Columns.Item(1).Cells.Borders.Item(ppBorderLeft).Weight = 3;
	table.Columns.Item(3).Cells.Borders.Item(ppBorderRight).Weight = 3;
	table.Columns.Item(7).Cells.Borders.Item(ppBorderLeft).Weight = 3;
	table.Columns.Item(9).Cells.Borders.Item(ppBorderRight).Weight = 3;
	
	table.Rows.Item(1).Cells.Borders.Item(ppBorderTop).Weight = 3;
	table.Rows.Item(3).Cells.Borders.Item(ppBorderBottom).Weight = 3;
	table.Rows.Item(7).Cells.Borders.Item(ppBorderTop).Weight = 3;
	table.Rows.Item(9).Cells.Borders.Item(ppBorderBottom).Weight = 3;
	
	return table;
}

function ulg_removeBorders(cell)
{
	cell.Borders.Item(ppBorderRight).Visible = 0;
	cell.Borders.Item(ppBorderBottom).Visible = 0;
}

function urg_removeBorders(cell)
{
	cell.Borders.Item(ppBorderLeft).Visible = 0;
	cell.Borders.Item(ppBorderBottom).Visible = 0;
}

function llg_removeBorders(cell)
{
	cell.Borders.Item(ppBorderTop).Visible = 0;
	cell.Borders.Item(ppBorderRight).Visible = 0;
}

function lrg_removeBorders(cell)
{
	cell.Borders.Item(ppBorderLeft).Visible = 0;
	cell.Borders.Item(ppBorderTop).Visible = 0;
}

function convertSudoku()
{
	function getCellNumber(x, y)
	{
		return tblCentralGrid.Cell(x + 1, y + 1).Shape.TextFrame.TextRange.Text.getFirstNumber();
	}
	
	for (var i = 0;i < 9;i++)
	{
		var tr = sudoku.insertRow();
		for (var j = 0;j < 9;j++)
			tr.insertCell().innerText = getCellNumber(i, j);
	}
}

function $convertCentralGrid()
{
	function getCellNumber(x, y)
	{
		return tblCentralGrid.Cell(x + 1, y + 1).Shape.TextFrame.TextRange.Text.getFirstNumber();
	}
	
	for (var i = 0;i < 9;i++)
		for (var j = 0;j < 9;j++)
		{
			var cell = $centralGrid.rows[i].cells[j];
			cell.innerText = getCellNumber(i, j);
			if (cell.relative)
				cell.relative.innerText = cell.innerText;
		}
}

function $convertGrid(table, grid)
{
	function getCellNumber(x, y)
	{
		return table.Cell(x + 1, y + 1).Shape.TextFrame.TextRange.Text.getFirstNumber();
	}
	
	for (var i = 0;i < 9;i++)
		for (var j = 0;j < 9;j++)
			if (!grid.rows[i].cells[j].relative)
				grid.rows[i].cells[j].innerText = getCellNumber(i, j);
}

function saveShapeAsImage(shape, path)
{
	var format;
	switch (FSO.GetExtensionName(path).toUpperCase())
	{
	case "EMF" :
		format = 5;
		break;
	case "GIF" :
		format = 0;
		break;
	case "JPG" :
		format = 1;
		break;
	case "PNG" :
		format = 2;
		break;
	case "WMF" :
		format = 4;
		break;
	default :
		format = 3;
	}
	
	shape.Export(path, format);
}

function setTextFormatting(table)
{
	for (var i = 1;i <= 9;i++)
		for (var j = 1;j <= 9;j++)
		{
			var tf = table.Cell(i, j).Shape.TextFrame;
			tf.HorizontalAnchor = 2;
			tf.VerticalAnchor = 3;
			tf.TextRange.Font.Size = 24;
		}
}

function $createSudokuTable(slide)
{
	var table = slide.Shapes.AddTable(21, 21, -1, -1, 624, 624).Table;
	
	table.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}");
	table.TableDirection = 1;
	
	table.Columns.Item(1).Cells.Borders.Item(ppBorderLeft).Weight = 3;
	table.Columns.Item(3).Cells.Borders.Item(ppBorderRight).Weight = 3;
	table.Columns.Item(7).Cells.Borders.Item(ppBorderLeft).Weight = 3;
	table.Columns.Item(9).Cells.Borders.Item(ppBorderRight).Weight = 3;
	table.Columns.Item(13).Cells.Borders.Item(ppBorderLeft).Weight = 3;
	table.Columns.Item(15).Cells.Borders.Item(ppBorderRight).Weight = 3;
	table.Columns.Item(19).Cells.Borders.Item(ppBorderLeft).Weight = 3;
	table.Columns.Item(21).Cells.Borders.Item(ppBorderRight).Weight = 3;
	
	table.Rows.Item(1).Cells.Borders.Item(ppBorderTop).Weight = 3;
	table.Rows.Item(3).Cells.Borders.Item(ppBorderBottom).Weight = 3;
	table.Rows.Item(7).Cells.Borders.Item(ppBorderTop).Weight = 3;
	table.Rows.Item(9).Cells.Borders.Item(ppBorderBottom).Weight = 3;
	table.Rows.Item(13).Cells.Borders.Item(ppBorderTop).Weight = 3;
	table.Rows.Item(15).Cells.Borders.Item(ppBorderBottom).Weight = 3;
	table.Rows.Item(19).Cells.Borders.Item(ppBorderTop).Weight = 3;
	table.Rows.Item(21).Cells.Borders.Item(ppBorderBottom).Weight = 3;
	
	var col = table.Columns.Item(10).Cells;
	for (var i = 1;i <= 6;i++)
		llg_removeBorders(col.Item(i));
	
	for (i = 16;i <= 21;i++)
		ulg_removeBorders(col.Item(i));
	
	col = table.Columns.Item(11).Cells;
	for (i = 1;i <= 6;i++)
		col.Item(i).Borders.Item(ppBorderTop).Visible = 0;
	
	for (i = 16;i <= 21;i++)
		col.Item(i).Borders.Item(ppBorderBottom).Visible = 0;
	
	col = table.Columns.Item(12).Cells;
	for (i = 1;i <= 6;i++)
		lrg_removeBorders(col.Item(i));
	
	for (i = 16;i <= 21;i++)
		urg_removeBorders(col.Item(i));
	
	col = table.Rows.Item(10).Cells;
	for (i = 1;i <= 6;i++)
		urg_removeBorders(col.Item(i));
	
	for (i = 16;i <= 21;i++)
		ulg_removeBorders(col.Item(i));
	
	col = table.Rows.Item(11).Cells;
	for (i = 1;i <= 6;i++)
		col.Item(i).Borders.Item(ppBorderLeft).Visible = 0;
	
	for (i = 16;i <= 21;i++)
		col.Item(i).Borders.Item(ppBorderRight).Visible = 0;
	
	col = table.Rows.Item(12).Cells;
	for (i = 1;i <= 6;i++)
		lrg_removeBorders(col.Item(i));
	
	for (i = 16;i <= 21;i++)
		llg_removeBorders(col.Item(i));
	
	return table;
}

//MISC. Functions
function markBlankCells()
{
	var blankCells = filterFullCells(document.getElementsByTagName("TD"));
	for (var i = 0;i < blankCells.length;i++)
		blankCells[i].wasBlank = true;
}

function askToClose(st)
{
	switch (st)
	{
	case ST_NORMAL :
		WScript.StdOut.WriteLine("Press Enter to quit.");
		/*
		WScript.StdOut.WriteLine("Your sudoku was solved. Now, ");
		WScript.StdOut.WriteLine("you can do one of the following:");
		WScript.StdOut.WriteLine("1. Type nothing and press Enter ");
		WScript.StdOut.WriteLine("to quit the app.");
		WScript.StdOut.WriteLine('2. Type "C" and press Enter to ');
		WScript.StdOut.WriteLine("copy the solved sudoku table into");
		WScript.StdOut.WriteLine(" clipboard.");
		WScript.StdOut.WriteLine("3. Enter the path of an image file");
		WScript.StdOut.WriteLine(" to save the sudoku table as an image.");
		WScript.StdOut.WriteLine("The following extensions are supported:");
		WScript.StdOut.WriteLine("*.BMP | *.EMF | *.GIF | *.JPG |");
		WScript.StdOut.WriteLine(" *.PNG | *.WMF");
		*/
		var input = WScript.StdIn.ReadLine();
		if (!input)
			WScript.Quit();
		
		WScript.StdOut.WriteLine("Generating sudoku table...");
		
		var PP = createPP();
		
		var prs = PP.Presentations.Add(0);
		
		var sld = prs.Slides.Add(1, 12);
		
		var tbl = createSudokuTable(sld);
		
		for (var i = 0;i < 9;i++)
			for (var j = 0;j < 9;j++)
			{
				var tf = tbl.Cell(i + 1, j + 1).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = sudoku.rows[i].cells[j];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 24;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
		
		if (input.toUpperCase() == "C")
		{
			tbl.Parent.Cut();
			WScript.StdOut.WriteLine("Sudoku table was successfully copied to clipboard.");
		}
		else
		{
			saveShapeAsImage(tbl.Parent, input);
			WScript.StdOut.WriteLine("Sudoku table was successfully saved as an image file.");
		}
		
		prs.Close();
		break;
	case ST_SAMURAI :
		WScript.StdOut.WriteLine("Press Enter to quit.");
		/*
		WScript.StdOut.WriteLine("Your sudoku is solved. Now, you can do ");
		WScript.StdOut.WriteLine("one of the following:");
		WScript.StdOut.WriteLine("1. Type nothing and press Enter to quit");
		WScript.StdOut.WriteLine(" the app.");
		WScript.StdOut.WriteLine('2. Type "C" and press Enter to copy the');
		WScript.StdOut.WriteLine(" solved sudoku table into clipboard.");
		WScript.StdOut.WriteLine("3. Enter the path of an image file to ");
		WScript.StdOut.WriteLine("save the sudoku table as an image.");
		WScript.StdOut.WriteLine("The following extensions are supported:");
		WScript.StdOut.WriteLine("*.BMP | *.EMF | *.GIF | *.JPG |");
		WScript.StdOut.WriteLine(" *.PNG | *.WMF");
		*/
		var input = WScript.StdIn.ReadLine();
		if (!input)
			WScript.Quit();
		
		WScript.StdOut.WriteLine("Generating sudoku table...");
		
		var PP = createPP();
		
		var prs = PP.Presentations.Add(0);
		
		var sld = prs.Slides.Add(1, 12);
		
		var tbl = $createSudokuTable(sld);
		
		for (var i = 1;i <= 6;i++)
			for (var j = 1;j <= 21;j++)
			{
				if ((9 < j) && (j < 13))
					continue;
				
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = j <= 9 ? $upperLeftGrid.rows[i - 1].cells[j - 1] : $upperRightGrid.rows[i - 1].cells[j - 13];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
		
		for (i = 7;i <= 9;i++)
		{
			for (j = 1;j <= 6;j++)
			{
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = $upperLeftGrid.rows[i - 1].cells[j - 1];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
			
			for (j = 7;j <= 15;j++)
			{
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = $centralGrid.rows[i - 7].cells[j - 7];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
			
			for (j = 16;j <= 21;j++)
			{
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = $upperRightGrid.rows[i - 1].cells[j - 13];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
		}
		
		for (i = 10;i <= 12;i++)
			for (j = 7;j <= 15;j++)
			{
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = $centralGrid.rows[i - 7].cells[j - 7];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
		
		for (i = 13;i <= 15;i++)
		{
			for (j = 1;j <= 6;j++)
			{
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = $lowerLeftGrid.rows[i - 13].cells[j - 1];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
			
			for (j = 7;j <= 15;j++)
			{
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = $centralGrid.rows[i - 7].cells[j - 7];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
			
			for (j = 16;j <= 21;j++)
			{
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = $lowerRightGrid.rows[i - 13].cells[j - 13];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
		}
		
		for (i = 16;i <= 21;i++)
			for (j = 1;j <= 21;j++)
			{
				if ((9 < j) && (j < 13))
					continue;
				
				var tf = tbl.Cell(i, j).Shape.TextFrame;
				tf.HorizontalAnchor = 2;
				tf.VerticalAnchor = 3;
				var c = j <= 9 ? $lowerLeftGrid.rows[i - 13].cells[j - 1] : $lowerRightGrid.rows[i - 13].cells[j - 13];
				tf.TextRange.Text = c.innerText;
				tf.TextRange.Font.Size = 18;
				if (c.wasBlank)
					tf.TextRange.Font.Color.RGB = RGB_GREEN;
			}
		
		if (input.toUpperCase() == "C")
		{
			tbl.Parent.Cut();
			WScript.StdOut.WriteLine("Sudoku table was successfully copied to clipboard.");
		}
		else
		{
			saveShapeAsImage(tbl.Parent, input);
			WScript.StdOut.WriteLine("Sudoku table was successfully saved as an image file.");
		}
		
		prs.Close();
		break;
	default :
		WScript.StdOut.Write("Press Enter to quit.");
		WScript.StdIn.Skip(1);
		WScript.Quit();
	}
}

function checkScript()
{
	if (FSO.GetBaseName(WScript.FullName).toLowerCase() != "cscript")
	{
		var wshShell = WScript.CreateObject("WScript.Shell");
		var cmdLine = 'cscript.exe "' + WScript.ScriptFullName + '"';
		var arg = (new Enumerator(WScript.Arguments)).item();
		if (arg)
			arg = arg.toUpperCase();
		if (arg == "/N")
			cmdLine += " /N";
		else if (arg == "/S")
			cmdLine += " /S";
		else if (arg == "/D")
			cmdLine += " /D";
		else if (arg == "/C")
			cmdLine += " /C";
		//Rerun this app by cscript.exe
		wshShell.Run(cmdLine);
		WScript.Quit();
	}
}

function checkArguments()
{
	if (WScript.Arguments.Count() > 1)
	{
		WScript.StdErr.WriteLine("Too many arguments.");
		askToClose();
	}
	
	if (WScript.Arguments.Count())
	{
		var arg = (new Enumerator(WScript.Arguments)).item().toUpperCase();
		if (arg == "/N")
			sudokuType = ST_NORMAL;
		else if (arg == "/S")
			sudokuType = ST_SAMURAI;
		else if (arg == "/C")
			sudokuType = ST_CLIPBOARD;
		else if (arg == "/D")
			sudokuType = ST_DESIGN;
		else
		{
			WScript.StdErr.WriteLine("Invalid arguments.");
			askToClose();
		}
	}
	else
	{
		WScript.StdOut.WriteLine("Which type of sudoku do you want to solve?");
		WScript.StdOut.WriteLine("1. Normal 9x9 Sudoku");
		WScript.StdOut.WriteLine("2. Samurai Sudoku");
		WScript.StdOut.WriteLine("You might also want to...");
		WScript.StdOut.WriteLine("3. Solve the sudoku stored in clipboard.");
		WScript.StdOut.WriteLine("4. Enter your unsolved sudoku in PowerPoint.");
		WScript.StdOut.Write("Please enter the number: ");
		var n = parseInt(WScript.StdIn.ReadLine());
		switch (n)
		{
		case 1 :
			sudokuType = ST_NORMAL;
			break;
		case 2 :
			sudokuType = ST_SAMURAI;
			break;
		case 3 :
			sudokuType = ST_CLIPBOARD;
			break;
		case 4 :
			sudokuType = ST_DESIGN;
			break;
		default :
			WScript.StdErr.WriteLine("Invalid number.");
			askToClose();
		}
	}
}

function isFull()
{
	var col = document.getElementsByTagName("TD");
	for (var i = 0;i < col.length;i++)
		if (!col[i].innerText)
			return false;
	return true;
}

function filterFullCells(cells)
{
	var result = []; //Array for storing blank cells.
	for (var i = 0;i < cells.length;i++)
		if (!cells[i].innerText)
			result.push(cells[i]);
	return(result);
}

function Doubt(number, range)
{
	this.invokeGuesses = Doubt_invokeGuesses;
	this.clearScope = Doubt_clearScope;
	this.number = number;
	this.range = range;
	this.scope = filterFullCells(document.getElementsByTagName("TD"));
	this.guesses = [];
	for (var i in range)
		this.guesses[i] = new Guess(this, i);
}

function Doubt_invokeGuesses()
{
	for (var i in this.guesses)
		if (this.guesses[i].invokeGuess())
			return true;
	return false;
}

function Doubt_clearScope()
{
	for (var i in this.scope)
		this.scope[i].innerText = "";
}

function Guess(parentDoubt, rangeItem)
{
	this.invokeGuess = Guess_invoke;
	this.parentDoubt = parentDoubt;
	this.rangeItem = rangeItem;
	this.result = RES_NOT_ACCOMPLISHED;
	this.subdoubt = null;
}

function Guess_invoke()
{
	var cell = this.parentDoubt.range[this.rangeItem];
	cell.innerText = this.parentDoubt.number;
	if (cell.relative)
		cell.relative.innerText = this.parentDoubt.number;
	
	if (!((sudokuType == ST_NORMAL) ? _tryToSolve() : $_tryToSolve()))
	{
		this.parentDoubt.clearScope();
		this.result = RES_FAILURE;
		return false;
	}
	
	if (isFull())
	{
		this.result = RES_SUCCESS;
		return true;
	}
	
	this.subdoubt = (sudokuType == ST_NORMAL) ? findDoubt() : $findDoubt();
	var r = this.subdoubt.invokeGuesses();
	this.result = r ? RES_SUCCESS : RES_FAILURE;
	return r;
}