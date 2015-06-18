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



## DataBase
Class `DataBase` consists of three member variables.

* `FileFullName` (String) is used to save the full name of BibTeX file.
* `LiteratureList` (ArrayList) is the set of literatures in BibTeX file.
* `Progress` (Double) is used to save the progress of loading database.



## Literature

## ErrorMessage
