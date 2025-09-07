import Foundation

let outputFile = "MYFILE.txt"

// Get current date and time in yyyy/MM/dd HH:mm:ss format
let now = Date()
let formatter = DateFormatter()
formatter.dateFormat = "yyyy/MM/dd HH:mm:ss"
let nowString = formatter.string(from: now)

let content = """
This program was written in Swift.
Current time is: \(nowString)
"""

do {
    try content.write(toFile: outputFile, atomically: true, encoding: .utf8)
    print("Successfully wrote current time to '\(outputFile)'.")
} catch {
    print("Error: Failed to write to file '\(outputFile)': \(error)")
    exit(1)
}
