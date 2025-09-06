package main

import (
	"fmt"
	"log"
	"os"
	"time"
)

// Main function to run the logging program.
func main() {
	const filename = "MYFILE.txt"

	// Create a log message with a timestamp.
	nowtime := time.Now().Format("2006/01/02 15:04:05")
	logMessage := fmt.Sprintf("This program is written in Go.\nCurrent Time = %s\n", nowtime)

	// Attempt to write the log message to the file.
	if err := writeLog(filename, logMessage); err != nil {
		log.Fatalf("Program terminated with an error: %v", err)
	}

	fmt.Printf("Successfully appended log to %s.\n", filename)
}

// writeLog appends a given message to a file, creating it if it doesn't exist.
func writeLog(filename string, message string) error {
	// Open the file in append mode. Create it if it doesn't exist.
	// Use read-write permissions 0644 (read/write for owner, read-only for others).
	file, err := os.OpenFile(filename, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer file.Close()

	// Write the message to the file.
	if _, err := fmt.Fprint(file, message); err != nil {
		return fmt.Errorf("failed to write to file: %w", err)
	}

	return nil
}
