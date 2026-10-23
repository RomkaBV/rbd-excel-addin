/* Public deployment settings only. Never put credentials or access tokens here. */
window.RBD_CONFIG = window.RBD_CONFIG || {
  version: "3.4.9",
  // Preserve the existing RBD calendar: Monday–Friday, 09:00–18:00 (9 hours).
  // For a lunch break use intervals: [["09:00","13:00"],["14:00","18:00"]].
  workCalendar: {timeZone:"Europe/Kyiv",weekdays:[1,2,3,4,5],intervals:[["09:00","18:00"]],holidays:[]},
  secureApi: {baseUrl:"",tenantId:"",clientId:"",scope:""}
};
