(async () => {
    const table = await getRelevantTable(PDFViewerApplication.pdfDocument);

    const excelOutput = table
        .map(row => `${row.name}\t${row.quantity}\t${row.price}`)
        .join("\n");

    console.log(excelOutput);

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

        let tableLines = relevantLines.slice(firstLine, lastLine + 1);

        const rowLengths = tableLines.map(line => line.length);
        const minRowWidth = Math.min(...rowLengths);
        const maxRowWidth = Math.max(...rowLengths);
        const optionalColumnCount = maxRowWidth - minRowWidth;

        let optionalColumnIndex = null;

        if (optionalColumnCount > 0) {
            optionalColumnIndex = Number(prompt(
                `${optionalColumnCount} column found which does not always have a value.\n` +
                "Please input the index of the missing column:\n\n" +
                tableLines
                    .toSorted((a, b) => b.length - a.length)
                    .map(line => `${line.map((value, index) => `${index}: ${value}`).join("\n")}`)
                    .join("\n\n")
            ));
        }

        tableLines = tableLines.map(line => {
            const row = [...line];

            // add an empty value where the optional column is missing
            if (row.length < maxRowWidth) row.splice(optionalColumnIndex, 0, "");

            return row;
        });

        const columnInput = prompt(
            "Enter the index of the columns which represent:\n" +
            "Product Name, Quantity, Total Price\n\n" +
            "eg: 1, 3, 4. \n" +
            "If one heading spans multiple columns separate with a /\n" +
            "eg: 0/1, 2/3, 4\n\n" +
            tableLines
                .map(line => `${line.map((value, index) => `${index}: ${value}`).join("\n")}`)
                .join("\n\n")
    );

        const [nameInput, quantityInput, priceInput] = columnInput
            .split(",")
            .map(column => column.trim());

        const columns = {
            name: nameInput
                .split("/")
                .map(column => Number(column.trim())),
            quantity: quantityInput
                .split("/")
                .map(column => Number(column.trim())),
            price: priceInput
        };

        return tableLines.map(line => {
            const row = [...line];

            return {
                name: columns.name
                    .map(column => row[column])
                    .join(" - ")
                    .trim(),
                quantity: columns.quantity
                    .map(column => Number(row[column]))
                    .reduce((total, value) => total * value, 1),
                price: Number(row[columns.price].replace(" 0", "").replace(/[^0-9.-]/g, ""))
            };
        });
    }
})();