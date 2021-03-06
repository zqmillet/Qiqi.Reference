# Namespace `_BibTeX`

This namespace provides three classes `DataBase`, `Literature` and `ErrorMessage`.

## Contents

* [1. Analysis of BibTeX Format](#AnalysisBibTeX)
* [2. Class `DataBase`](#ClassDataBase)
* [3. Class `Literature`](#ClassLiterature)
* [4. Class `ErrorMessage`](#ClassErrorMessage)


<h2 id="AnalysisBibTeX">Analysis of BibTeX Format</h2>
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

The second format exists ambiguity sometimes. For instance:
	
	@Article{BibTeXKey, 
		Author = "Author1, Author2 and Author3",  
		Title  = "This is the "title" of literature", 
		Year   = "2015"
	}

There are four quotations in the third line, we can interpret this line in two different ways:

1. Title is **This is the**.
2. Title is **This is the "title" of literatue**.

Which is correct is decided by the rest of the code. If we interpret this line in the first way, the rest of the code has syntax error. So the second way is correct. In another words, we must analyse all the code to decide how to interpret the quotations.

There is another problem. The BibTeX from [ScienceDirect](http://www.sciencedirect.com/) is shown in following code.

	@Article{BibTeXKey, 
		Author   = "Author1, Author2 and Author3",  
		Title    = "This is the "title" of literature", 
		Year     = "2015",
		Keywords = "Keyword1",
		Keywords = "Keyword2",
		Keywords = "Keyword3" 
	}

[JabRef](http://jabref.sourceforge.net/) handles this kind of BibTeX with a BUG. It only adds the first keyword into the property `Keywords`.

Therefore, I wrote the following three classes `DataBase`, `Literature` and `ErrorMessage`.

<h2 id="ClassDataBase">Class <code>DataBase</code></h2>
Class `DataBase` consists of three member variables.

* `FileFullName` (String) is used to save the full name of BibTeX file.
* `LiteratureList` (ArrayList) is the set of literatures in BibTeX file.
* `Progress` (Double) is used to save the progress of loading database.

Class `DataBase` has a constructor with parameter.

	Public Sub New(ByVal FileFullName As String)

When a `DataBase` is created, it only save the file name of database, and initialize the literature list. When the sub `DataBaseLoading` is called, this class will read and analyse the database.

If the sub `DataBaseLoading` find the char `@`, it begins to read the type of the literature. When the sub `DataBaseLoading` reader the first char `{`, it will begin to read the BibTeXKey of the literature. When it reach the first char `,`, it will create a new literature.

When the progress is 10%, 20%, ... , 100%, the class `DataBase` will raise a event `ProgressUpdate`. This event is used to show the progress of database loading.

<h2 id="ClassLiterature">Class <code>Literature</code></h2>
Class `Literature` consists of three member variables.

* `Type` (String) is used to save the type of the literature.
* `ID` (String) is used to save the BibTeXKey of the literature.
* `PropertyList` (ArrayList) is the set of properties of the literature.

Class `Literature` has a constructor with parameter.

	Public Sub New(ByVal Type As String, _
				   ByVal ID As String, _
				   ByVal LiteratureBuffer As String, _
				   ByRef ErrorMessage As _BibTeX.ErrorMessage)
         
If the class `DataBase` have the type and the BibTeXKey of a literature, it will create a new literature with this constructor. The `Type` is the type of the literature, the `ID` is the BibTeXKey of the literature, the `LiteratureBuffer` is used to save the rest of the literature. For example, the whole code of a literature is shown as following code.

	@Article{BibTeXKey, 
		Author = "Author1, Author2 and Author3",  
		Title  = "This is the "title" of literature", 
		Year   = "2015"
	}

The `Type` is `Article`, the `ID` is `BibTeXKey`, the `LiteratureBuffer` is shown as following code.

		Author = "Author1, Author2 and Author3",  
		Title  = "This is the "title" of literature", 
		Year   = "2015"

The `ErrorMessage` is used to record the error message of lexical analysis.

At last, the constructor calls the function `LexicalAnalysis` to analyse and add the properties of the literature into `PropertyList`.

The definition of function `LexicalAnalysis` is shown in following code.

	 Private Function LexicalAnalysis(ByVal LiteratureBuffer As String, _
									  ByRef ErrorMessage As _BibTeX.ErrorMessage) _
									  As Boolean

This function anaylses the `LiteratureBuffer`, then add the properties into the `PropertyList`. If there is no error in `LiteratureBuffer`, the returned value is `True`, else `False`. The error message is stored in the `ErrorMessage` if there is error in `LiteratureBuffer`.

When this function read a quotation, it will enter the `QuotationMode`. In the `QuotationMode`, if it read another quotation, it firstly assumes that this quotation is the end char of the value. Then the function will continue to process the rest code. If there is syntax error in the rest code, the function knows the assumption is wrong, so it regards that quotation is a char of the value, and find another quotation. Recursion is used to realize this process.

<h2 id="ClassErrorMessage">Class <code>ErrorMessage</code></h2>
Class `ErrorMessage` consists of three member variables.

* `ExistError` (Boolean) represent whether there exist error.
* `Message` (String) is used to save the message of the error.
* `LineNumber` (Integer) is used to save the line number of the error.

To avoid that user make mistakes, the menber variables `ExistError` and `Message` are private. Therefore, class provide four public functions/subs to operate these two member variables.
