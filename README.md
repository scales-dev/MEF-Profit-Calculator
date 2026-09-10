# Profit Calculator

---

From our invoice pdf's we would like to know how much profit we would make assuming we sell everything we bought on our website.
Also, if they change their prices, we would like to see that change, so we can changes ours aswell.

The idea of this, is to open an invoice from any of our suppliers, paste in the script in script.js, and get a spreadsheet.
The issue is, in order to track price changes I would either need a server, or have it run on the same computer every time.
To track profit I will need to have up to date online shop prices as well.
Since our suppliers name their products vastly different to our website, 
eg. "granadilla" meaning "a box of 20 granadilla",
We need a spreadsheet to match the names to quantities/weights, depending on how we sell it.
Because every supplier formats their invoice differently the pdf parsing has to be robust, which is scary.
What originally sounded like a small simple project is shaping up to be pretty big...

## How it works
- Open a supply invoice PDF in browser
  - Why? I sadly chose what I think is the best language for the job, which is very sadly JavaScript
    - Most people have firefox installed, unlike other interpreters such as the JVM or wtv python uses
    - When downloading a PDF from an email it opens it in the browser anyway
    - It is very easy to get the text from a PDF with JavaScript on firefox
    - It is easy to use the script, just press "ctrl+shift+i" click console and paste the script
- [Read PDF](https://github.com/scales-dev/MEF-Profit-Calculator/blob/master/parse_invoice.js)
  - Get supplier name 
  - Find the lines representing the table of what we bought
    - Determine which columns are "product name", "quantity" and "individual price"
- [Compare prices](https://github.com/scales-dev/MEF-Profit-Calculator/blob/master/form_spreadsheet.ts)
  - Get the actual products we bought as their names would be on the website 
    - Get this from a spreadsheet which maps the names the supplier calls them, to the name and quantity our website uses
  - Output a CSV table with each product and expected profit
    - Keep track of price change history
    - Highlight high price markups


I hope that explains simply how it works. 
Decomposition is key when starting a project, it helps me a lot to write down an exact plan to follow.