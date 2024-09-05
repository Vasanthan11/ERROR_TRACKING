document.getElementById('extractButton').addEventListener('click', extractComments);

document.getElementById('fileInput').addEventListener('change', function(event) {
    var fileCount = event.target.files.length;
    var fileCountText = fileCount > 0 ? fileCount + ' file(s) chosen' : 'No files chosen';
    document.getElementById('fileCount').textContent = fileCountText;

    // Set the upload date when files are selected
    if (fileCount > 0) {
        const uploadDate = new Date().toLocaleDateString();
        document.getElementById('uploadDate').value = uploadDate;
    }
});

async function extractComments() {
    const fileInput = document.getElementById('fileInput').files;
    if (fileInput.length === 0) {
        alert('Please upload at least one PDF file.');
        return;
    }

    const uploadDate = document.getElementById('uploadDate').value || new Date().toLocaleDateString();

    let comments = [];
    for (const file of fileInput) {
        const pdfData = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;

        // Extract banner name, week, PRF number, and page from the filename
        const fileName = file.name;

        // Updated regex patterns for parsing filename
        const bannerNameMatch = fileName.match(/^(.+?)_/); // Extract everything before the first underscore
        const weekMatch = fileName.match(/W[Kk](\d+)/); // Extract week number after "WK" or "Wk"
        const prfNumberMatch = fileName.match(/PRF(\d+)/); // Extract PRF number after "PRF"
        const pageMatch = fileName.match(/_P(\d+|[A-Z]\d+)/); // Extract page number after "_P", allowing letters

        const bannerName = bannerNameMatch ? bannerNameMatch[1] : 'Unknown';
        const week = weekMatch ? weekMatch[1] : 'Unknown';
        const prfNumber = prfNumberMatch ? prfNumberMatch[1] : 'Unknown';
        const page = pageMatch ? pageMatch[1] : 'Unknown';

        for (let i = 0; i < pdf.numPages; i++) {
            const pdfPage = await pdf.getPage(i + 1);
            const annotations = await pdfPage.getAnnotations();

            annotations.forEach(annotation => {
                if (annotation.subtype !== 'Popup') {
                    // Determine error type based on the comment's content
                    let errorType = 'Product_Description'; // Default category
                    let content = annotation.contents || 'No content';

                    if (content.toLowerCase().includes('price')) {
                        errorType = 'Price_Point';
                    } else if (content.toLowerCase().includes('alignment')) {
                        errorType = 'Overall_Layout';
                    } else if (content.toLowerCase().includes('image')) {
                        errorType = 'Image_Usage';
                    }

                    let errorsContent = '';
                    let gdContent = '';

                    // If the content starts with 'GD:', separate it out
                    if (content.startsWith('GD:')) {
                        gdContent = content.substring(3).trim(); // Remove 'GD:' and trim the content
                    } else {
                        errorsContent = content;
                    }

                    // Collect comments with GD content separated
                    comments.push({
                        Date: uploadDate,
                        BannerName: bannerName,
                        Week: week,
                        PRFNumber: prfNumber,
                        FileName: file.name,
                        Errors: errorsContent,
                        QC: annotation.title || 'Unknown',
                        GD: gdContent, // Place GD content in the GD column
                        ErrorType: errorType,
                       
                    });
                }
            });
        }

        // Add a blank row after each PDF's comments
        comments.push({
            BannerName: '',
            Week: '',
            PRFNumber: '',
            Page: '',
            FileName: '',
            Errors: '',
            GD: '',
            QC: '',
            ErrorType: '',
            Date: ''
        });
    }

    exportToExcel(comments);
}

function exportToExcel(comments) {
    // Define the worksheet and workbook
    const worksheet = XLSX.utils.json_to_sheet(comments, { header: ["Date","BannerName", "Week", "PRFNumber", "FileName", "Errors","QC", "GD",  "ErrorType", ] });
    const workbook = XLSX.utils.book_new();

    // Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Comments");

    // Create the Excel file and trigger a download
    XLSX.writeFile(workbook, "comments.xlsx");
}
