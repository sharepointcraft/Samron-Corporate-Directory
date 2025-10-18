import * as React from "react";
import { useState, useEffect } from "react";
import styles from './CorporateDirectory.module.scss';
import type { ICorporateDirectoryProps } from './ICorporateDirectoryProps';
import "@fortawesome/fontawesome-free/css/all.min.css";
import { HttpClient } from "@microsoft/sp-http";

interface IPerson {
  FullName: string;
  Department: string;
  PhoneNumber: string;
  SecondaryPhone?: string;
  Email?: string;
  EXT?: string;
  Location: string;
  Initials?: string;
  ProfileColor?: string;
  [key: string]: any;
}

const CorporateDirectory: React.FC<ICorporateDirectoryProps> = ({ context, documentLibrary, csvFile }) => {
  const [people, setPeople] = useState<IPerson[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [filteredPeople, setFilteredPeople] = useState<IPerson[]>([]);
  const [activeFilter, setActiveFilter] = useState<{ type: string; value: string }[]>([{ type: "letter", value: "All" }]);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 9;

  // Styles for print table

  const cardColors = [
    "#000", "#000", "#000", "#000", "#000", "#000", "#000", "#000", "#000",
  ];

  function parseCSVRow(row: string): string[] {
    const cols: string[] = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < row.length; i++) {
      const char = row[i];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === "," && !inQuotes) {
        cols.push(cur);
        cur = "";
      } else {
        cur += char;
      }
    }
    cols.push(cur);
    return cols.map((c) => c.trim());
  }

  const fetchPeople = async (): Promise<void> => {
    try {
      const decodedLibraryName = decodeURIComponent(documentLibrary);
      const decodedFileName = decodeURIComponent(csvFile);
      const fileUrl = `${context.pageContext.web.absoluteUrl}/${decodedLibraryName}/${decodedFileName}`;
      const response = await context.httpClient.get(fileUrl, HttpClient.configurations.v1);
      const raw = await response.text();
      const rows = raw
        .split(/\r?\n/)
        .filter((line, idx) => idx > 0 && line.trim().length > 0);

      const parsedPeople: IPerson[] = rows.map((row) => {
        const cols = parseCSVRow(row);
        const lastName = cols[0] || "";
        const firstName = cols[1] || "";
        const ext = cols[2] || "";
        const department = cols[3] || "";
        const primaryPhone = cols[4] || "";
        const secondaryPhone = cols[5] || "";
        const email = cols[6] || "";

        const fullName = `${firstName} ${lastName}`.trim();
        const initials = `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase();

        return {
          FullName: fullName,
          PhoneNumber: primaryPhone,
          SecondaryPhone: secondaryPhone,
          Email: email,
          EXT: ext,
          Initials: initials,
          Department: department,
          Location: "",
          ProfileColor: "", // Optional, you may assign colors here
        };
      });

      setPeople(parsedPeople.filter((p) => p.FullName));
    } catch (error) {
      console.error("Error loading CSV data:", error);
    }
  };

  useEffect(() => {
    fetchPeople();
  }, [documentLibrary, csvFile]);

  useEffect(() => {
    let filtered = [...people];
    activeFilter.forEach((filter) => {
      if (filter.type === "letter" && filter.value !== "All") {
        filtered = filtered.filter((person) =>
          person.FullName?.toLowerCase().startsWith(filter.value.toLowerCase())
        );
      } else if (filter.type === "Department" && filter.value !== "All") {
        filtered = filtered.filter((person) => person.Department === filter.value);
      }
    });

    if (searchTerm) {
      const lowerSearch = searchTerm.toLowerCase();
      const fullNameMatches = filtered.filter(person =>
        person.FullName?.toLowerCase().includes(lowerSearch)
      );
      const otherMatches = filtered.filter(person =>
        (
          person.Department?.toLowerCase().includes(lowerSearch) ||
          person.EXT?.toLowerCase().includes(lowerSearch) ||
          person.PhoneNumber?.toLowerCase().includes(lowerSearch) ||
          person.SecondaryPhone?.toLowerCase().includes(lowerSearch)
        ) &&
        !fullNameMatches.includes(person)
      );
      filtered = [...fullNameMatches, ...otherMatches];
    }
    setFilteredPeople(filtered);
    setCurrentPage(1);
  }, [people, activeFilter, searchTerm]);

  const handleAlphabetFilterChange = (type: string, value: string) => {
    setSearchTerm("");
    if (value === "All") {
      setActiveFilter([{ type: "letter", value: "All" }]);
    } else {
      setActiveFilter((prevFilters) => {
        const letterFilterIndex = prevFilters.findIndex((filter) => filter.type === "letter");
        if (letterFilterIndex !== -1) {
          const newFilters = [...prevFilters];
          newFilters[letterFilterIndex] = { type, value };
          return newFilters;
        } else {
          return [...prevFilters, { type, value }];
        }
      });
    }
  };

  const handleDepartmentFilterChange = (type: string, value: string) => {
    setSearchTerm("");
    setActiveFilter((prevFilters) => {
      const DepartmentFilterIndex = prevFilters.findIndex((filter) => filter.type === "Department");
      if (DepartmentFilterIndex !== -1) {
        const newFilters = [...prevFilters];
        newFilters[DepartmentFilterIndex] = { type, value };
        return newFilters;
      } else {
        return [...prevFilters, { type, value }];
      }
    });
  };

  const totalPages = Math.ceil(filteredPeople.length / itemsPerPage);
  const paginatedData = filteredPeople.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  const Departments: string[] = [
    ...Array.from(new Set(people.map((p) => p.Department?.trim()).filter((title) => title))),
  ].sort((a, b) => a.localeCompare(b));

  return (
    <>
      <div className={styles.directoryTitle}>
        Employee Directory
        <div className={styles.instructionTooltip}>
          <i className="fas fa-info-circle" style={{ marginLeft: '8px', fontSize: '16px', color: '#666' }}></i>
          <div className={styles.tooltipContent}>
            <strong>Instruction:</strong><br />
            You must first edit the web part and fill in the property pane fields with the following values:<br /><br />
            <strong>Library Internal Name</strong> - The internal name of the SharePoint document library where the CSV file is uploaded.<br /><br />
            <strong>CSV File Name</strong> - The exact name of the uploaded CSV file (including the file extension, e.g., sampleData.csv).
          </div>
        </div>
      </div>

      <div className={styles.directoryContainer}>
        <aside className={styles.sidebar}>
          <FilterSection
            title="Department"
            options={Departments}
            type="Department"
            activeFilter={activeFilter.find((f) => f.type === "Department")}
            handleFilterChange={handleDepartmentFilterChange} />
        </aside>

        <main className={styles.mainContent}>
          <div className={styles.exportData}>
            <div className={styles.searchContainer}>  
              <input
                type="text"
                placeholder="Search People..."
                className={styles.searchBox}
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
              {searchTerm && (
                <span
                  className={styles.clearSearch}
                  onClick={() => setSearchTerm("")}
                  title="Clear search"
                >&times;</span>
              )}
            </div>
            <button className={styles.exportButton} onClick={() => printFilteredDirectory(filteredPeople)}>Print</button>

          </div>

          <AlphabetFilter
            activeFilter={activeFilter.find((f) => f.type === "letter")}
            handleFilterChange={handleAlphabetFilterChange} />

          <div className={styles.peopleGrid}>
            {paginatedData.map((person, index) => (
              <PersonCard
                key={index}
                person={{
                  ...person,
                  ProfileColor: cardColors[index % cardColors.length]
                }}
              />
            ))}
          </div>

          {paginatedData.length === 0 && <div className={styles.noResult}>No result found</div>}

          <Pagination
            totalPages={totalPages}
            currentPage={currentPage}
            onPageChange={setCurrentPage} />
        </main>
      </div>     
    </>
  );
};

const printFilteredDirectory = (filteredPeople: IPerson[]) => {
  if (!filteredPeople || filteredPeople.length === 0) {
    alert("No records to print.");
    return;
  }

  //Group users by Department
  const groupedByDept = filteredPeople.reduce((acc, person) => {
    const dept = person.Department?.trim() || "No Department";
    if (!acc[dept]) acc[dept] = [];
    acc[dept].push(person);
    return acc;
  }, {} as Record<string, IPerson[]>);

  //Sort departments A–Z (No Department last)
  const sortedDepartments = Object.keys(groupedByDept).sort((a, b) => {
    if (a === "No Department") return 1;
    if (b === "No Department") return -1;
    return a.localeCompare(b);
  });

  // Sort users by FullName A–Z within each department
  sortedDepartments.forEach((dept) => {
    groupedByDept[dept].sort((a, b) =>
      (a.FullName || "").localeCompare(b.FullName || "")
    );
  });

  // Build HTML content
  const printContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>Office Directory</title>
<style>
    @page {
      size: A4 portrait;
      margin: 10mm;

      /* Hide browser-added header/footer info */
      @top-left { content: none }
      @top-center { content: none }
      @top-right { content: none }
      @bottom-left { content: none }
      @bottom-center { content: none }
      @bottom-right { content: none }
    }

  * {
    box-sizing: border-box;
  }
  html, body {
    margin: 0;
    padding: 0;
    height: auto;
    color: #000;
    font-family: Arial, sans-serif;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  h2 {
    text-align: center;
    text-transform: uppercase;
    font-size: 18px;
    margin: 10px 0 15px 0;
    font-weight: bold;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 0;
    table-layout: fixed;
  }
  th, td {
    border: 1px solid #000;
    padding: 6px 8px;
    font-size: 11.5px;
    text-align: left;
    vertical-align: middle;
    word-wrap: break-word;
  }
  th {
    background-color: #000 !important;
    color: #fff !important;
    font-weight: bold;
    border-bottom: 0.2px solid #fff !important;
  }
  .dept-header td {
    background-color: #000 !important;
    color: #fff !important;
    font-weight: bold;
    font-size: 13px;
    border: 1px solid #000;
    text-transform: capitalize;
  }

  th:first-child, td:first-child { width: 11%; }  /* FIRST */
  th:nth-child(2), td:nth-child(2) { width: 11%; } /* LAST */
  th:nth-child(3), td:nth-child(3) { width: 6%; }  /* EXT */
  th:nth-child(4), td:nth-child(4) { width: 24%; } /* DEPARTMENT */
  th:nth-child(5), td:nth-child(5) { width: 13%; } /* COMPANY PHONE */
  th:nth-child(6), td:nth-child(6) { width: 13%; } /* CELL PHONE */
  th:nth-child(7), td:nth-child(7) { width: 22%; } /* EMAIL */

  @media print {
    html, body {
      height: auto !important;
      overflow: visible !important;
    }
    table, tr, td, th {
      page-break-inside: avoid !important;
    }
    h2 {
      page-break-after: avoid !important;
    }
    thead {
      display: table-header-group;
    }
  }
</style>
      </head>
      <body>
        <div id="printArea">
          <h2>OFFICE DIRECTORY</h2>
          <table>
            <thead>
              <tr>
                <th>FIRST</th>
                <th>LAST</th>
                <th>EXT</th>
                <th>TITLE</th>
                <th>COMPANY PHONE</th>
                <th>CELL PHONE</th>
                <th>EMAIL</th>
              </tr>
            </thead>
            <tbody>
              ${sortedDepartments
                .map(
                  (dept) => `
                    <tr class="dept-header">
                      <td colspan="7">${dept}</td>
                    </tr>
                    ${groupedByDept[dept]
                      .map((p) => {
                        const [firstName = "", lastName = ""] =
                          p.FullName?.split(" ") || [];
                        return `
                          <tr>
                            <td>${firstName}</td>
                            <td>${lastName}</td>
                            <td>${p.EXT || ""}</td>
                            <td>${p.Department || ""}</td>
                            <td>${p.PhoneNumber || ""}</td>
                            <td>${p.SecondaryPhone || ""}</td>
                            <td>${p.Email || ""}</td>
                          </tr>`;
                      })
                      .join("")}
                  `
                )
                .join("")}
            </tbody>
          </table>
        </div>
      </body>
    </html>
  `;

  //Use a hidden iframe — no new browser window
  const printFrame = document.createElement("iframe");
  printFrame.style.position = "fixed";
  printFrame.style.width = "0";
  printFrame.style.height = "0";
  printFrame.style.border = "none";
  document.body.appendChild(printFrame);

  const doc = printFrame.contentWindow?.document;
  if (doc) {
    doc.open();
    doc.write(printContent);
    doc.close();

    //Wait a bit for rendering to avoid page break between title and table
    setTimeout(() => {
      const win = printFrame.contentWindow;
      if (win) {
        win.focus();
        win.print();
      }
    }, 500); // Slight delay ensures full layout render
  }

  //Clean up the iframe after printing
  setTimeout(() => document.body.removeChild(printFrame), 1500);
};






const truncateText = (text: string, maxLength: number) => {
  if (!text) return "";
  return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
};

const PersonCard: React.FC<{ person: IPerson }> = ({ person }) => (
  <div className={styles.personCard}>
    <div className={styles.initials} style={{ backgroundColor: person.ProfileColor }}>
      {person.Initials}
    </div>

    <div className={styles.personDetails}>
      <h3>{person.FullName}</h3>

      {person.Department && (
        <p className={styles.singleLine}>{truncateText(person.Department, 30)}</p>
      )}

      {person.EXT && (
        <p className={styles.singleLine}>EXT: {truncateText(person.EXT, 30)}</p>
      )}

      {person.PhoneNumber && (
        <p className={styles.singleLine}>
          <i className="fas fa-phone" style={{ marginRight: '6px' }}></i>
          <a href={`tel:${person.PhoneNumber}`} className={styles.phoneLink}>{person.PhoneNumber}</a>
        </p>
      )}

      {person.SecondaryPhone && person.SecondaryPhone !== person.PhoneNumber && (
        <p className={styles.singleLine}>
          <i className="fas fa-phone" style={{ marginRight: '6px' }}></i>
          <a href={`tel:${person.SecondaryPhone}`} className={styles.phoneLink}>{person.SecondaryPhone}</a>
        </p>
      )}

      {person.Email && (
        <p className={styles.singleLine}>
          <i className="fas fa-envelope" style={{ marginRight: "6px" }}></i>
          <a href={`mailto:${person.Email}`} className={styles.phoneLink}>{person.Email}</a>
        </p>
      )}
    </div>
  </div>
);

interface FilterSectionProps {
  title: string;
  options: string[];
  type: string;
  activeFilter: { type: string; value: string } | undefined;
  handleFilterChange: (type: string, value: string) => void;
}

const FilterSection: React.FC<FilterSectionProps> = ({
  title,
  options,
  type,
  activeFilter,
  handleFilterChange,
}) => {
  const isActive = activeFilter?.type === type;

  const handleClear = () => {
    handleFilterChange(type, "All");
  };

  const handleButtonClick = (option: string) => {
    handleFilterChange(type, option);
  };

  return (
    <div className={styles.filterBox}>
      <h3>{title}</h3>
      {isActive && activeFilter?.value !== "All" && (
        <button className={styles.clearButton} onClick={handleClear}>
          Clear
        </button>
      )}
      <div className={styles.filterButtons}>
        {options.map((option) => (
          <button
            key={option}
            onClick={() => handleButtonClick(option)}
            className={
              isActive && activeFilter?.value === option
                ? `${styles.filterButton} ${styles.active}`
                : styles.filterButton
            }
          >
            {option}
          </button>
        ))}
      </div>
    </div>
  );
};

interface AlphabetFilterProps {
  activeFilter: { type: string; value: string } | undefined;
  handleFilterChange: (type: string, value: string) => void;
}

const AlphabetFilter: React.FC<AlphabetFilterProps> = ({
  activeFilter,
  handleFilterChange,
}) => {
  const alphabet = ["All", ..."ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("")];
  return (
    <div className={styles.alphabetFilter}>
      {alphabet.map((letter) => (
        <button
          key={letter}
          onClick={() => handleFilterChange("letter", letter)}
          className={activeFilter?.type === "letter" && activeFilter?.value === letter ? styles.active : ""}
        >
          {letter}
        </button>
      ))}
    </div>
  );
};

interface PaginationProps {
  totalPages: number;
  currentPage: number;
  onPageChange: (page: number) => void;
}

const Pagination: React.FC<PaginationProps> = ({
  totalPages,
  currentPage,
  onPageChange,
}) => {
  const generatePageNumbers = () => {
    const pages: (number | string)[] = [];

    pages.push(1);

    if (currentPage > 4) {
      pages.push("...");
    }

    const startPage = Math.max(2, currentPage - 2);
    const endPage = Math.min(totalPages - 1, currentPage + 2);

    for (let i = startPage; i <= endPage; i++) {
      if (i !== 1 && i !== totalPages) {
        pages.push(i);
      }
    }

    if (currentPage < totalPages - 3) {
      pages.push("...");
    }

    if (totalPages > 1) {
      pages.push(totalPages);
    }

    return pages;
  };

  const pageNumbers = generatePageNumbers();

  return (
    <div
      className={styles.pagination}
      style={{ display: totalPages > 0 ? "flex" : "none" }}
    >
      {currentPage > 1 && <button onClick={() => onPageChange(1)}>&laquo;</button>}
      {currentPage > 1 && <button onClick={() => onPageChange(currentPage - 1)}>&lt;</button>}

      {pageNumbers.map((page, index) =>
        page === "..." ? (
          <span key={`ellipsis-${index}`} className={styles.ellipsis}>
            ...
          </span>
        ) : (
          <button
            key={page}
            onClick={() => onPageChange(Number(page))}
            className={currentPage === page ? styles.active : ""}
          >
            {page}
          </button>
        )
      )}

      {currentPage < totalPages && (
        <button onClick={() => onPageChange(currentPage + 1)}>&gt;</button>
      )}
      {currentPage < totalPages && <button onClick={() => onPageChange(totalPages)}>&raquo;</button>}
    </div>
  );
};

export { CorporateDirectory };