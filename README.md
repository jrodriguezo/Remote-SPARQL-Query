# Remote-SPARQL-Query :mag:

Application to make SPARQL queries to ontologies dataset serialized in RDF through a Fuseki endpoint.

## Neccesary software 

* Python3 or superior </br>
  https://www.python.org/downloads/
* Apache Jena Fuseki </br>
  https://jena.apache.org/download/index.cgi

## Installation 

1. Download repository </br> 
`git clone https://github.com/jrodriguezo/Remote-SPARQL-Query`

2. Indicate the fuseki directory on line 70 </br> 
`os.chdir(os.getcwd()+"\\fuseki")      # C:\Users\[User]\Desktop\APP\fuseki`

3. Install the libraries </br> 
`py -m pip install xlsxwriter` </br> 
`py -m pip install SPARQLWrapper` </br> 
_NOTE: Tkinter is included by default in Python._

## Interactions

1. Open application where it was saved through command line </br>
`C:\Users\[User]\Desktop\APP> py "Remote SPARQL Query.py"` </br> 
_NOTE: Another way is double click on the file._

2. Load ontology dataset </br> 
_Choose file > select ontology file path_ </br> 
_NOTE: Ontology must be in RDF or turtle format._

3. SPARQL queries </br> 
_Introduce SPARQL syntax > search button_

4. Save ontology in Excel </br> 
_Save > select path where you want to save_


