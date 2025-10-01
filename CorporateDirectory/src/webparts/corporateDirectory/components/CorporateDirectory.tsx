import * as React from "react";
import { useState, useEffect } from "react";
import styles from './CorporateDirectory.module.scss';
import type { ICorporateDirectoryProps } from './ICorporateDirectoryProps';
import "@fortawesome/fontawesome-free/css/all.min.css";
import { HttpClient } from "@microsoft/sp-http";

interface IPerson {
  FullName: string;
  JobTitle: string;
  Department: string;
  PhoneNumber: string;
  SecondaryPhone?: string;
  Email?: string;
  Supervisor?: string;
  Location: string;
  Initials?: string;
  ProfileColor?: string; // Keep this, but we'll assign it differently
  [key: string]: any;
}

const CorporateDirectory: React.FC<ICorporateDirectoryProps> = ({ context, documentLibrary, csvFile }) => {

  const [people, setPeople] = useState<IPerson[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [filteredPeople, setFilteredPeople] = useState<IPerson[]>([]);
  const [activeFilter, setActiveFilter] = useState<{ type: string; value: string }[]>(
    [{ type: "letter", value: "All" }]
  );
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 9;

  // Define your set of unique colors for the cards
  const cardColors = [
    
    "#2ECC71", // Emerald Green
    "#4169E1", // Royal Blue
    "#E94E77", // Raspberry
    "#F4D03F", // Sunflower Yellow
    "#8E44AD", // Amethyst
    "#E67E22", // Carrot Orange
    "#008080", // Teal
    "#C0392B", // Pomegranate
    "#00BCD4", // Cyan
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
      const libraryName = documentLibrary
      const fileName = csvFile

      const decodedLibraryName = decodeURIComponent(libraryName);
      const decodedFileName = decodeURIComponent(fileName);
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
        //const rawDisplayName = cols[2] || "";
        const jobTitle = cols[3] || "";
        const rawSupervisor = cols[2] || "";
        const primaryPhone = cols[4] || "";
        const secondaryPhone = cols[5] || "";
        const email = cols[6] || "";

        const fullName = `${firstName} ${lastName}`.trim();
        let supervisor = rawSupervisor;
        if (supervisor.includes(",")) {
          const [supLast, supFirst] = supervisor.split(",").map((s) => s.trim());
          supervisor = `${supLast} ${supFirst}`;
        }

        const initials = `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase();

        return {
          FullName: fullName,
          JobTitle: jobTitle,
          PhoneNumber: primaryPhone,
          SecondaryPhone: secondaryPhone,
          Email: email,
          Supervisor: supervisor,
          Initials: initials,
          Department: "",
          Location: "",
          // ProfileColor will be assigned at rendering time for paginated data
        };
      });
        setPeople(parsedPeople.filter((p) => p.FullName));   
            //const filteredPeople = parsedPeople.filter((p) => p.FullName);
            //setPeople(filteredPeople.slice(0, 15));
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
      } else if (filter.type === "JobTitle" && filter.value !== "All") {
        filtered = filtered.filter((person) => person.JobTitle === filter.value);
      }
    });

    if (searchTerm) {
      const lowerSearch = searchTerm.toLowerCase();

      const fullNameMatches = filtered.filter(person =>
        person.FullName?.toLowerCase().includes(lowerSearch)
      );

      const otherMatches = filtered.filter(person =>
        (
          person.JobTitle?.toLowerCase().includes(lowerSearch) ||
          person.Supervisor?.toLowerCase().includes(lowerSearch) ||
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

  const handleJobTitleFilterChange = (type: string, value: string) => {
    setSearchTerm("");
    setActiveFilter((prevFilters) => {
      const jobTitleFilterIndex = prevFilters.findIndex((filter) => filter.type === "JobTitle");
      if (jobTitleFilterIndex !== -1) {
        const newFilters = [...prevFilters];
        newFilters[jobTitleFilterIndex] = { type, value };
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

  const jobTitles: string[] = [
    ...Array.from(new Set(people.map((p) => p.JobTitle?.trim()).filter((title) => title))),
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
            options={jobTitles}
            type="JobTitle"
            activeFilter={activeFilter.find((f) => f.type === "JobTitle")}
            handleFilterChange={handleJobTitleFilterChange} />
        </aside>

        <main className={styles.mainContent}>
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
              >
                &times;
              </span>
            )}
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
                  ProfileColor: cardColors[index % cardColors.length] // Assign color here
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
      </div></>
  );
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

      {person.JobTitle && ( // Changed from person.Supervisor to person.JobTitle for display
        <p className={styles.singleLine}>{truncateText(person.JobTitle, 30)}</p>
      )}

      {person.Supervisor && (
        <p className={styles.singleLine}>EXT: {truncateText(person.Supervisor, 30)}</p>
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

export default CorporateDirectory;