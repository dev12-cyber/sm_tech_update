// Ensure menu loads after DOM content is fully loaded
document.addEventListener("DOMContentLoaded", function () {

    // Define the Excel file path
    const excelFilePath = "assets/sheets/services.xlsx";

    // Function to generate the services menu
    function generateServicesMenu(services) {
        const menuContainer = document.getElementById("services-menu");
        if (!menuContainer) {
            console.error("❌ services-menu not found in this page!");
            return;
        }

        menuContainer.innerHTML = ""; // Clear previous content

        services.forEach(service => {
            const serviceID = encodeURIComponent(service.ID.replace(/ /g, "_"));
            const serviceHTML = `<li><a href="service-details.html?id=${encodeURIComponent(serviceID)}">${service.ID}</a></li>`;
            menuContainer.innerHTML += serviceHTML; // Append dynamically
        });
    }

    // Fetch and process the Excel file
    function fetchExcelFile(filePath) {
        fetch(filePath)
            .then(response => response.arrayBuffer())
            .then(data => {
                const workbook = XLSX.read(data, { type: "array" });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet);
                generateServicesMenu(rows);
            })
            .catch(error => console.error("❌ Error fetching Excel file:", error));
    }

    // Load services menu
    fetchExcelFile(excelFilePath);
});
