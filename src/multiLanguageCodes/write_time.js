const fs = require('fs').promises; // Use the promise-based version of the fs module

/**
 * @brief Appends a timestamped log message to a file.
 *
 * @param {string} filename The path to the file.
 * @param {string} message The message to append to the file.
 * @returns {Promise<void>} A promise that resolves when the operation is complete.
 */
const writeLog = async (filename, message) => {
    try {
        await fs.appendFile(filename, message);
        console.log(`Successfully appended log to '${filename}'.`);
    } catch (err) {
        console.error('An error occurred while writing to the file:', err);
        throw err; // Propagate the error to the caller
    }
};

/**
 * @brief Returns the current time formatted as a string.
 * @returns {string} The formatted time string (e.g., '2023/10/27 15:30:00').
 */
const getFormattedCurrentTime = () => {
    const now = new Date();
    const pad = (num) => String(num).padStart(2, '0');

    const year = now.getFullYear();
    const month = pad(now.getMonth() + 1);
    const day = pad(now.getDate());
    const hours = pad(now.getHours());
    const minutes = pad(now.getMinutes());
    const seconds = pad(now.getSeconds());

    return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
};

// ----------------------------------------------------
// Main execution
// ----------------------------------------------------
const main = async () => {
    const filename = 'MYFILE.txt';
    const logMessage = `This program is written in JavaScript.\nCurrent Time = ${getFormattedCurrentTime()}\n`;

    try {
        await writeLog(filename, logMessage);
    } catch (err) {
        // The error is already logged inside writeLog, so we can handle it gracefully here.
        console.error('Program terminated with an error.');
        process.exit(1);
    }
};

main();
