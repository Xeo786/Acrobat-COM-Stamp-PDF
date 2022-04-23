# Acrobat-COM-Stamp-PDF
It can Stamp thousands of pdf in short time 

# Requirements
1) Acrobat DC pro Installed, we need its COM libs so can't work without it
2) Do not think It can find text of scanned pdf (images in pdf) page be smart and understand the possibility of capability

# How to use
Select folder and It will place stamp on every PDF found in the folder,

# How it works
1) find text on page if found the proceed further
2) get total word from `.getPageNumWords`
3) check text one by one `.objJSO.getPageNthWord(pagenumber, a_index)` amazingly its very very fast do not know why
4) select text with `.SetTextSelection()` and I didn't know it was whole different com object `"AcroExch.HiliteList"`
5) get selection coordniates by `.GetBoundingRect`
6) Place signature right round those coords

# Example
```autohotkey
findtext := "Chop"
FileSelectFolder, signfolder,,,Select folder
signfolder := signfolder "\*.pdf"
Loop, files, %signfolder%, R
	PDFtext(A_LoopFileFullPath,findtext)
msgbox, all pdf have been signed
return
```
# Custom Stamps
1) goto Menu >> View >> Comments >> Annotation
2) there you find Stamp
3) Create a Custom Stamp https://acrobatusers.com/tutorials/stamps-acrobat-x/
4) Stamp this pdf on some document
5) now goto Menu >> View >> Tools >> Javascript
6) Open Java script debugger
7) now write this JS into console `this.selectedAnnots[0].AP`
8) now select place stamp once again using mouse
9) return to console click inside console at the End of above JS Press Enter
10) It console will print something like `#ZUZY4X5AO4_F_ewyPMFYmC`
11) now you will ise this stamp address to stamp pdf

```autohotkey
	CustomStampScript =
(
var annot = this.addAnnot({
	page: %pagenum%,
	type: "Stamp",
	author: "Xeo786",
	name: "Arbab",
	rect: [%left%, %top%, %right%, %bottom%],
	contents: "",
	AP: "#ZUZY4X5AO4_F_ewyPMFYmC"
	});
)
```
