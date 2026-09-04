(async () => {
    console.log(getRelevantTable(PDFViewerApplication.pdfDocument));

    async function getRelevantTable(pdf) {
        let lines = [];
        let lineContent = [];

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();

            for (const i of content.items) {
                const string = i.str.trim();
                if (string) lineContent.push(string);

                if (i.hasEOL) {
                    lines.push(lineContent);
                    lineContent = [];
                }
            }
        }

        const relevantLines = lines.filter(line => line.length > 1);

        const firstLine = Number(prompt(
            "What is the first line containing product data?\n\n" +
            relevantLines.map((line, i) => `${i}: ${line.join(" | ")}`).join("\n")
        ));

        const lastLine = Number(prompt(
            "What is the last line containing product data?\n\n" +
            relevantLines.slice(firstLine+1).map((line, i) => `${i+firstLine+1}: ${line.join(" | ")}`).join("\n")
        ));

        const categoryTitles = relevantLines.at(firstLine-1).map((line, i) => `${i}: ${line}`);

        const columnInput = prompt(
            "Enter the index of the columns which represent the following:\n" +
            "Product Name, Quantity, Individual Price\n\n" +
            categoryTitles.join("\n") +
            "\n\neg: 1, 3, 4"
        );

        const columns = columnInput
            .split(",")
            .map(column => Number(column.trim()));

        const tableLines = relevantLines.slice(firstLine, lastLine);
        tableLines.sort((a, b) => a.length - b.length);

        const minRowWidth = tableLines.at(0).length;
        const maxRowWidth = tableLines.at(tableLines.length-1).length;
        const optionalColumns = maxRowWidth - minRowWidth;

        let optionalColumnIndex = 0;
        if (optionalColumns) optionalColumnIndex = Number(
            prompt(
                `${optionalColumns} column found which does not always have a value.\n` +
                "Please input the index of the column in the list:\n" +
                `${categoryTitles.filter((title, i) => !columns.includes(i)).join("\n")}`
            )
        );

        return relevantLines
            .slice(firstLine, lastLine)
            .map(line => {
                // add an empty value where the optional column is missing
                if (line.length < maxRowWidth) {
                    line.splice(optionalColumnIndex, 0, "");
                }

                // then get the three selected columns, thats all we care for
                return columns.map(column => line[column]);
            });
    }
})();