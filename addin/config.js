/* Default configuration: "byok" (bring your own Claude API key), used when the
   add-in is hosted as static files (GitHub Pages). On assistant.bulwarkhq.com.au
   the Worker serves its own /config.js with mode "service" (subscription). */
window.DBN_CONFIG=window.DBN_CONFIG||{mode:'byok'};
