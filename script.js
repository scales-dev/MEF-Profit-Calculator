(async () => {
    const table = await getRelevantTable(PDFViewerApplication.pdfDocument);
    console.log(table);

    async function getRelevantTable(pdf) {
        const lines = [];
        let line = [];

        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
            const page = await pdf.getPage(pageNumber);
            const content = await page.getTextContent();

            for (const item of content.items) {
                const text = item.str.trim();

                if (text) line.push(text);
                if (item.hasEOL) {
                    lines.push(line);
                    line = [];
                }
            }
        }

        const relevantLines = lines.filter(line => line.length > 1);

        const lineText = relevantLines
            .map((line, index) => `${index}: ${line.join(" | ")}`)
            .join("\n");

        const firstLine = Number(prompt(
            "What is the first line containing product data?\n\n" +
            lineText
        ));

        const remainingLines = relevantLines
            .slice(firstLine + 1)
            .map((line, index) => `${index + firstLine + 1}: ${line.join(" | ")}`)
            .join("\n");

        const lastLine = Number(prompt(
            "What is the last line containing product data?\n\n" +
            remainingLines
        ));

        const tableLines = relevantLines.slice(firstLine, lastLine + 1);

        const columnTitles = relevantLines[firstLine - 1];

        const columnInput = prompt(
            "Enter the index of the columns which represent:\n" +
            "Product Name, Quantity, Individual Price\n\n" +
            columnTitles
                .map((title, index) => `${index}: ${title}`)
                .join("\n") +
            "\n\neg: 1, 3, 4"
        );

        const columns = columnInput
            .split(",")
            .map(column => Number(column.trim()));

        const rowLengths = tableLines.map(line => line.length);
        const minRowWidth = Math.min(...rowLengths);
        const maxRowWidth = Math.max(...rowLengths);

        const optionalColumnCount = maxRowWidth - minRowWidth;

        let optionalColumnIndex = null;

        if (optionalColumnCount > 0) {
            optionalColumnIndex = Number(prompt(
                `${optionalColumnCount} column found which does not always have a value.\n` +
                "Please input the index of the column in the list:\n" +
                columnTitles
                    .filter((_, index) => !columns.includes(index))
                    .map((title, index) => `${index}: ${title}`)
                    .join("\n")
            ));
        }

        return tableLines.map(line => {
            const row = [...line];

            // add an empty value where the optional column is missing
            if (row.length < maxRowWidth) row.splice(optionalColumnIndex, 0, "");

            // then get the three selected columns, thats all we care for
            return columns.map(column => row[column]);
        });
    }
})();