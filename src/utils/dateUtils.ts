export function formatSharePointDate(value: unknown): string {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  // Handle Excel serial date (e.g., 45979)
  // Excel base date is Dec 30, 1899.
  // Formula: (Serial - 25569) * 86400 * 1000 gives milliseconds since Unix epoch (Jan 1 1970)
  // However, JS Date uses 1970 as epoch.
  // A simpler way often cited is new Date((serial - (25567 + 2)) * 86400 * 1000)
  // Let's use the standard conversion:
  if (typeof value === 'number' || (typeof value === 'string' && !isNaN(Number(value)) && !value.includes('-') && !value.includes('/'))) {
    const serial = Number(value);
    // 25569 is the offset between Excel (1900-01-01) and Unix (1970-01-01)
    // Adjust for leap year bug in Excel (1900 was not a leap year) if necessary, 
    // but usually standard offset works for modern dates.
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;                                      
    const date_info = new Date(utc_value * 1000);

    // Correction: The above gives UTC. We want to display it as is.
    // Often simply: new Date(Math.round((serial - 25569)*86400*1000))
    const date = new Date(Math.round((serial - 25569)*86400*1000));
    
    // Add timezone offset to keep the date "local" as represented in Excel
    // Or simpler: just extract day, month, year from the UTC date object?
    // Let's try standard JS date formatting.
    return date.toLocaleDateString("it-IT", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    });
  }

  // Handle ISO string or standard Date object
  const date = new Date(String(value));
  if (isNaN(date.getTime())) {
    return String(value); // Fallback if parsing fails
  }

  return date.toLocaleDateString("it-IT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  });
}
