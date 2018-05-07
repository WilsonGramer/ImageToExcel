//
//  main.swift
//  ImageToExcel
//
//  Created by Wilson Gramer on 5/5/18.
//  Copyright © 2018 Neef.co. All rights reserved.
//

import Foundation
import QuartzCore
import Cocoa

// MARK: - Helper stucts

/// From https://stackoverflow.com/a/30593673/5569234
extension Collection {
    /// Returns the element at the specified index iff it is within bounds, otherwise nil.
    subscript (safe index: Index) -> Element? {
        return indices.contains(index) ? self[index] : nil
    }
}

/// Stores the RGBA values of a pixel color.
struct Pixel {
    var r: CGFloat!
    var g: CGFloat!
    var b: CGFloat!
    var a: CGFloat!
}

// MARK: - Helper functions

/// Parses the path string into a file URL
func getURL(for pathString: String) -> URL {
    return URL(fileURLWithFileSystemRepresentation: pathString, isDirectory: false, relativeTo: nil)
}

/// Converts a `CGImage` to an array of `Pixel`s.
func getPixels(from image: CGImage) -> [Pixel] {
    let pixelData = image.dataProvider!.data!
    let data: UnsafePointer<UInt8> = CFDataGetBytePtr(pixelData) // Converts the image's pixel data to raw data
    let buffer = UnsafeBufferPointer(start: data, count: (4 * image.width * image.height)) // Multipy by 4 because there are 4 (R, G, B, and A) bytes per pixel
    let bufferArray = Array(buffer) // Converts the buffer to an array
    
    var pixels = [Pixel]()
    for (index, _) in bufferArray.enumerated() { // Iterates over the raw data, converting it to `Pixel`s
        if index == bufferArray.count { break }
        var pixel = Pixel()
        if index != 0 && index % 4 == 0 { // Only iterate over every 4th byte, since each pixel has 4 bytes
            pixel.r = CGFloat(bufferArray[index - 4])
            pixel.g = CGFloat(bufferArray[index - 3])
            pixel.b = CGFloat(bufferArray[index - 2])
            pixel.a = CGFloat(bufferArray[index - 1])
            pixels.append(pixel)
            //print("\(index) == (r: \(pixel.r!), g: \(pixel.g!), b: \(pixel.b!))")
        }
    }
    
    let suffix = Array(bufferArray.suffix(4))
    pixels.append(Pixel(
        r: CGFloat(suffix[0]),
        g: CGFloat(suffix[1]),
        b: CGFloat(suffix[2]),
        a: CGFloat(suffix[3])
    )) // Hack to get the last pixel of the image
    
    return pixels
}

/// Converts a hex string to a hex number
func getColor(hex: String) -> UInt64 {
    let scanner = Scanner(string: hex)
    scanner.scanLocation = 0
    
    var rgbValue: UInt64 = 0
    scanner.scanHexInt64(&rgbValue)
    
    return rgbValue
}

// MARK: - Main program

func main(_ inputPath: String, _ outputPath: String) {
    print("🔸 Input Path: \(inputPath)")
    print("🔸 Output Path: \(outputPath)")
    
    print("➡️ Reading image... ", terminator: "")
    
    let image = NSImage(contentsOfFile: inputPath)
    if image == nil { print("ERROR: Invalid path/image. Try again with a different path/image."); return }
    
    print("👍")
    
    print("➡️ Extracting pixels from image... ", terminator: "")
    
    let cgImage = image!.cgImage(forProposedRect: nil, context: nil, hints: nil)! // Converts the NSImage to a CGImage with default settings
    let pixels = getPixels(from: cgImage) // Gets the pixels from the cgImage into an array of `Pixel`s
    
    print("👍 (\(cgImage.width)x\(cgImage.height) image, \(pixels.count) total pixels)")
    
    ////////////////
    
    print("➡️ Creating spreadsheet... ", terminator: "")
    
    // Creates the spreadsheet
    let workbook: UnsafeMutablePointer<lxw_workbook>! = workbook_new(outputPath)
    let worksheet: UnsafeMutablePointer<lxw_worksheet>! = workbook_add_worksheet(workbook, nil)
    
    worksheet_set_column(worksheet, 0, lxw_col_t(cgImage.width), 2, nil) // Make the cells square so they look more like pixels
    
    for y in 1...(cgImage.height) { // Iterates over the image's height
        for x in 1...(cgImage.width) { // Iterates over the image's width
            let i = (y * cgImage.width + x) - 1 // Gets the index of the `Pixel` at the current width and height
            if i >= pixels.count { break } // Don't keep iterating if there aren't any more pixels!
            let p = pixels[i] // Gets the pixel at the index
            
            let hexString = String(format:"%02X", Int(p.r)) + String(format:"%02X", Int(p.g)) + String(format:"%02X", Int(p.b)) // Converts the pixel's color values into a hex string
            let color: UInt64 = getColor(hex: hexString) // Converts the hex string into a hex number
            
            ////////////////
            
            let format: UnsafeMutablePointer<lxw_format>! = workbook_add_format(workbook) // Creates the cell format
            
            // Sets the color of the cell's background, foreground, and font colors to the pixel's color
            format_set_pattern(format, 1) // Solid color
            format_set_bg_color(format, lxw_color_t(color))
            format_set_fg_color(format, lxw_color_t(color))
            format_set_font_color(format, lxw_color_t(color))
            
            worksheet_write_string(worksheet, lxw_row_t(y - 1), lxw_col_t(x - 1), "", format) // Writes the cell to the speadsheet
        }
    }
    
    print("👍")
    
    print("➡️ Saving spreadsheet... ", terminator: "")
    
    workbook_close(workbook) // Closes and saves the speadsheet
    
    print("👍")
    
    print("✅ Done! The spreadsheet is located at \(outputPath.replacingOccurrences(of: " ", with: "\\ ")).")
}

// MARK: - Initialize program (parse arguments)

if let inputPath = CommandLine.arguments[safe: 1] {
    let inputURL = getURL(for: inputPath) // Parses the input path string into a file URL
    
    var outputURL: URL
    
    // If the output path is specified, use that. Otherwise, create a file named "imageName.imageExtension.xslx" in the same directory
    if let outputPath = CommandLine.arguments[safe: 2] {
        outputURL = getURL(for: outputPath) // Parses the output path string into a file URL
    } else {
        outputURL = inputURL.appendingPathExtension("xlsx")
    }
    
    main(inputURL.path, outputURL.path)
} else {
    print("ERROR: Please specify a path to the input image.")
}
