# Namespace `_BibTeX`

This namespace provides three classes `DataBase`, `Literature` and `ErrorMessage`.

## Analysis of BibTeX Format
There are two sorts of formats of BibTeX.

The values of properties are marked with `{` and `}`, which is shown as following code.

	@Article{BibTeXKey, 
		Author = {Author1, Author2 and Author3},  
		Title  = {This is the title of literature}, 
		Year   = {2015}
	}

The values of properties are marked with `"`, which is shown as following code.

	@Article{BibTeXKey, 
		Author = "Author1, Author2 and Author3",  
		Title  = "This is the title of literature", 
		Year   = "2015"
	}

The second format exist ambiguity sometimes. For instance:
	
	@Article{BibTeXKey, 
		Author = "Author1, Author2 and Author3",  
		Title  = "This is the "title" of literature", 
		Year   = "2015"
	}

There are four quotations in the third line, we can interpret this line in two different ways:

1. Title is **This is the**
2. Title is **This is the "title" of literatue**

Which is correct is decided by the rest of the code. If we interpret this line in the first way, the rest of the code has syntax error. So the second way is correct. In another words, we must analyse all the code to decide how to interpret the quotations.

There is another problem. The BibTeX from [ScienceDirect](ScienceDirect.com) 

## DataBase
Class `DataBase` consists of three member variables.

* `FileFullName` (String) is used to save the full name of BibTeX file.
* `LiteratureList` (ArrayList) is the set of literatures in BibTeX file.
* `Progress` (Double) is used to save the progress of loading database.


## Literature



## ErrorMessage
