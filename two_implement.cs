 //****************************************************************************************************
 //
 //     2      Skriv programmet
 //
 //     I ett lagersystem behövs en funktion för att packa en lista med artiklar i ett antal lådor på ett speciellt sätt.
 //     Antalet lådor att packa i ska vara valbart och man vet att antalet artiklar att packa är mycket större än antalet lådor.
 //     Målet är att vikten ska fördelas så jämt som möjligt mellan lådorna. Beskriv ditt förslag på algoritm och skriv en klass BoxPacker
 //     (i valfritt programmeringsspråk, dock ej pseudokod) som implementerar den publika metoden Pack() enligt:
 //     public List<Box> Pack(List<Article> articles, int numBoxes)
 //
 //     Du får bygga vidare på klasserna Box och Article om du vill. Koden ska kompilera och vara så pass komplett att den skulle kunna användas i ett produktionssystem.
 //
 //**************************************************************************************************** 
 

class Box
{
    List<Article> BoxItems;
}

class Article
{
    int WeightInGrams; // interval is 100g – 1kg
}