// Drishti settings
// Paste the Web App URL from Apps Script between the quotes below, then save.
// It looks like https://script.google.com/macros/s/..../exec
// If you leave it empty, Drishti keeps data in each person's own browser only.
window.DRISHTI_API = "https://script.google.com/macros/s/AKfycbyS6jIwVM0gIhMHS9woiCdDqBQUQWsujZaM2X1h9d1KZIvfMwIwBqhXMNhJEpvhMAgu/exec";
// OneDrive: the Application (client) ID of the "Purview Drishti" app registration.
// It looks like 1a2b3c4d-1111-2222-3333-444455556666. Leave it empty and OneDrive stays switched off.
window.DRISHTI_MS_CLIENT_ID = "d07dcce6-f049-4e4f-bf8b-0695d895240b";
// OneDrive: the Microsoft account everyone signs in with. Every upload comes from My files in its OneDrive.
window.DRISHTI_MS_ACCOUNT = "Purviewautomate@purview-healthcare.com";
// OneDrive: leave this empty. Only if browsing should start in one folder instead of My files, paste that
// folder's link here (Share, then People with existing access, then Copy link).
window.DRISHTI_ONEDRIVE_FOLDER = "";
