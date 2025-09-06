import Foundation

/**
 * An enumeration representing possible errors during file operations.
 */
enum WriteTimeError: Error {
    case fileWriteFailed(String)
}

/**
 * A class to write a timestamped message to a file.
 * The file is automatically created if it does not exist.
 */
class WriteTime {
    
    private let filename: String
    
    init(filename: String) {
        self.filename = filename
    }
    
    /**
     * Appends a given message to the file.
     *
     * @param message The string message to be written.
     * @return A `Result` indicating success or failure with an associated error.
     */
    func write(message: String) -> Result<Void, WriteTimeError> {
        let fileURL = URL(fileURLWithPath: filename)
        
        do {
            try message.append(toFile: fileURL.path, atomically: true, encoding: .utf8)
            print("Successfully appended log to \(filename).")
            return .success(())
        } catch {
            let errorMessage = "Failed to write to file: \(error.localizedDescription)"
            print(errorMessage)
            return .failure(.fileWriteFailed(errorMessage))
        }
    }
}

// ----------------------------------------------------
// Main Execution
// ----------------------------------------------------

// Instantiate the WriteTime with the desired filename
let filename = "MYFILE.txt"
let writer = WriteTime(filename: filename)

// Get the current date and time and format it
let now = Date()
let formatter = DateFormatter()
formatter.dateFormat = "yyyy/MM/dd HH:mm:ss"
let nowtime = formatter.string(from: now)

// Create the log message using a multi-line string literal
let logMessage = """
This program is written in Swift.
Current Time = \(nowtime)

"""

// Write the log message to the file using the Result type
let result = writer.write(message: logMessage)

// Handle the result
switch result {
case .success():
    // The success case is already handled by the print statement inside the write method.
    break
case .failure(let error):
    // The failure case is also handled by the print statement.
    // For more complex applications, you might log the error here.
    print("Execution failed: \(error)")
}
