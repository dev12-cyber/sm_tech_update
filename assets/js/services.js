// Define the Excel file path (Ensure it is accessible via an HTTP server)
const excelFilePath = "assets/sheets/services.xlsx"; // Adjust as needed

// Function to fetch and process the Excel file
function fetchExcelFile(filePath, callback) {
    fetch(filePath)
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: "array" });

            // Get the first sheet and convert to JSON
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet);

            callback(rows); // Pass the fetched data to the callback function
        })
        .catch(error => console.error("Error fetching Excel file:", error));
}

// ✅ Function to generate dynamic service cards on the homepage
function generateHomepageServices(services) {
    const container = document.getElementById("services-container");
    if (!container) return;

    container.innerHTML = ""; // Clear existing content

    services.forEach(service => {
        const serviceID = encodeURIComponent(service.ID.replace(/ /g, "_"));

        const serviceHTML = `
            <div class="col-xl-3 col-lg-4 col-md-6 wow fadeInUp" data-wow-delay=".2s">
                <div class="service-box-items style-3">
                    <div class="icon">
                       <img src="${service.Icon}" alt="Service Icon" style="width: 40px; height: 40px;">

                    </div>
                    <div class="content">
                        <h3><a href="service-details.html?id=${encodeURIComponent(serviceID)}">${service.ID}</a></h3>
                        <p>${service.Home_Content}</p>
                        <div class="service-btn">
                            <a href="service-details.html?id=${encodeURIComponent(serviceID)}" class="arrow-icon">
                                <img src="assets/img/icon/02.svg" alt="Read More">
                            </a>
                            <a href="service-details.html?id=${encodeURIComponent(serviceID)}" class="link-btn">Read more</a>
                        </div>
                    </div>
                </div>
            </div>
        `;
        container.innerHTML += serviceHTML; // Append dynamically
    });
}

// ✅ Function to generate service menu (Sidebar or Navbar)
function generateServicesMenu(services) {
    const menuContainer = document.getElementById("services-menu");
    if (!menuContainer) return;

    menuContainer.innerHTML = ""; // Clear existing content

    services.forEach(service => {
        const serviceID = encodeURIComponent(service.ID.replace(/ /g, "_"));
        const serviceHTML = `<li><a href="service-details.html?id=${encodeURIComponent(serviceID)}">${service.ID}</a></li>`;
        menuContainer.innerHTML += serviceHTML; // Append dynamically
    });
}

// ✅ Function to get the service ID from the URL
function getServiceIdFromURL() {
    const params = new URLSearchParams(window.location.search);
    return params.get("id") ? params.get("id").replace(/_/g, " ") : null;
}

// ✅ Function to update an element if it exists
function updateElement(id, value) {
    const element = document.getElementById(id);
    if (element) {
        element.innerText = value;
    }
}

function loadServiceDetails() {
    const serviceId = getServiceIdFromURL();
    const detailsContainer = document.getElementById("service-details-container");

    if (!serviceId) {
        if (detailsContainer) detailsContainer.innerHTML = "<p>Service not found</p>";
        return;
    }

    fetchExcelFile(excelFilePath, (rows) => {
        generateServicesMenu(rows); // Load side menu

        const service = rows.find(row => row.ID?.trim() === serviceId);

        if (!service) {
            if (detailsContainer) detailsContainer.innerHTML = "<p>Service not found</p>";
            return;
        }

        // ✅ Populate existing elements
        updateElement("service-title", service.ID);
        updateElement("service-li", service.ID);
        updateElement("service-text", service.Home_Content);
        updateElement("service-title-1", service.ID);
        updateElement("service-title-2", service.ID);
        updateElement("service-title-3", service.ID);
        updateElement("service-title-4", service.ID);
        updateElement("service-title-5", service.ID);
        updateElement("service-title-6", service.ID);

        // ✅ Add new content from Excel columns
        updateElement("hero-content", service.Hero_content);
        updateElement("feature-title", service.Feature_title);
        updateElement("feature-content", service.Feature_content);
        updateElement("why-choose-title", service.Why_choose_title);
        updateElement("why-choose-content", service.Why_choose_content);
        updateElement("why-choose-content2", service.Why_choose_content2);
        updateElement("why-choose-title1", service.Why_choose_title);
        updateElement("why-choose-content1", service.Why_choose_content);
        updateElement("why-choose-content21", service.Why_choose_content2);
        updateElement("category", service.Link);

        // ✅ Update Image (Image1)
        if (service.Image1) {
            const imageElement = document.getElementById("service-image");
            if (imageElement) imageElement.src = service.Image1;
        }

        // ✅ Create clickable link for linked sheet
        if (service.Link) {
            const linkContainer = document.getElementById("linked-service");
            if (linkContainer) {
                linkContainer.innerHTML = `<a href="#" onclick="loadLinkedSheet('${service.Link}')">${service.Link.replace(/_/g, ' ')}</a>`;
            }
        }
    });
}

// ✅ Function to populate the marquee dynamically
function populateMarquee(services) {
    const titleElements = document.querySelectorAll(".service-title");

    if (titleElements.length === 0) return;

    titleElements.forEach((element, index) => {
        const service = services[index % services.length]; // Loop through services
        element.textContent = service.ID; // Set service title dynamically
    });
}

// ✅ Page-specific execution logic
document.addEventListener("DOMContentLoaded", function () {
    if (document.getElementById("services-menu")) {
        // Load services only if the services container is present (homepage)
        fetchExcelFile(excelFilePath, generateServicesMenu);
    }
    if (document.getElementById("services-container")) {
        // Load services only if the services container is present (homepage)
        fetchExcelFile(excelFilePath, (rows) => {
            generateHomepageServices(rows);
            generateServicesMenu(rows);
            populateMarquee(rows);
        });
    }

    if (document.getElementById("service-details-container")) {
        // Load service details only if on service details page
        loadServiceDetails();
    }
});


// ✅ Utility: Load workbook
function loadWorkbook(callback) {
    fetch(excelFilePath)
        .then(res => res.arrayBuffer())
        .then(buffer => {
            const workbook = XLSX.read(buffer, { type: 'array' });
            callback(workbook);
        })
        .catch(err => console.error('Error loading Excel:', err));
}

// ✅ Main function: Load IDs from all linked sheets and inject into accordion
function loadAllSheetIDsIntoAccordion(workbook) {
    const mainSheet = workbook.Sheets[workbook.SheetNames[0]];
    const linkRows = XLSX.utils.sheet_to_json(mainSheet);
    const container = document.querySelector("#accordion-container");
    if (!container) return;

    // Clear existing content
    container.innerHTML = "";

    // Iterate over each Link entry
    const queryString = window.location.search;
    const urlParams = new URLSearchParams(queryString);
    const sheetName = urlParams.get('id');
    if (!sheetName || !workbook.Sheets[sheetName]) return;

    const sheet = workbook.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(sheet);

    sheetData.forEach(sheetRow => {
        const id = sheetRow.ID;
        if (!id) return;

        const accordionHTML = `
            <div class="sidebar__toggle col-4" id="${id}">
                <div class="accordion-item mb-4 wow fadeInUp" data-wow-delay=".2s">
                    <h5 class="accordion-header">
                        <button class="accordion-button collapsed" type="button" style="background-color: #101011 !important;">
                            ${id.replace(/_/g, ' ')}
                        </button>
                    </h5>
                </div>
            </div>
        `;
        container.insertAdjacentHTML("beforeend", accordionHTML);

        // Attach the click event for each generated accordion item
        document.getElementById(id).addEventListener("click", function () {
            // Open the offcanvas
            $("#offcanvas2").addClass("info-open");
            $(".offcanvas__overlay").addClass("overlay-open");
            
            // Dynamically update offcanvas content with values from the selected row
            const offcanvasContent = document.querySelector(".offcanvas__content");
            const imageURL = `assets/img/service/${sheetName}/${sheetRow.ID}.png`;
            // Clear previous content
            offcanvasContent.innerHTML = `
                <div class="offcanvas__top mb-3 d-flex justify-content-between align-items-center">
                    <div class="offcanvas__logo">
                        <h4 class="accordion-header">
                            ${sheetRow.Tabs}
                        </h4>
                    </div>
                    <div class="offcanvas__close">
                        <button onclick="toggleOffcanvas('offcanvas1')">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                </div>
                <div class="breadcrumb-sub-title">
                    <img src="${imageURL}" alt="img" style="max-width: 100% !important; height">
                </div>
                <div class="offcanvas__details mb-3">
                    <p class="text d-none d-xl-block">${sheetRow.Tabs_Content || 'N/A'}</p>
                </div>
            `;
        });
    });
}

// ✅ Trigger on load (already exists)
if (document.getElementById("accordion-container")) {
    loadWorkbook(loadAllSheetIDsIntoAccordion);
}
