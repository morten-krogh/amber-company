// Render the balance report JSON produced by generate_balance_report.py.
// Usage: swift -module-cache-path /tmp/amber-swift-cache scripts/render_balance_report.swift INPUT.json OUTPUT.pdf
import Foundation
import AppKit
import CoreGraphics

struct Metadata: Decodable {
    let FNAMN: String
    let ORGNR: String
    let start: String
    let end: String
    let latest_voucher: String
    let generated: String
    let sha256: String
}
struct Row: Decodable {
    let kind: String
    let label: String
    let account: String
    let values: [String]?
}
struct Report: Decodable {
    let metadata: Metadata
    let rows: [Row]
    let notes: [String]
}

guard CommandLine.arguments.count == 3 else {
    fatalError("Usage: render_balance_report.swift INPUT.json OUTPUT.pdf")
}
let report = try JSONDecoder().decode(Report.self, from: Data(contentsOf: URL(fileURLWithPath: CommandLine.arguments[1])))
let output = URL(fileURLWithPath: CommandLine.arguments[2])
var page = CGRect(x: 0, y: 0, width: 595.28, height: 841.89)
guard let consumer = CGDataConsumer(url: output as CFURL),
      let context = CGContext(consumer: consumer, mediaBox: &page, nil) else {
    fatalError("Could not create PDF")
}
let formatter = NumberFormatter()
formatter.locale = Locale(identifier: "sv_SE")
formatter.numberStyle = .decimal
formatter.minimumFractionDigits = 2
formatter.maximumFractionDigits = 2
let columns: [CGFloat] = [24, 64, 337, 415, 493]
let widths: [CGFloat] = [35, 268, 75, 75, 78]
var y: CGFloat = 0
var pageNumber = 0

func text(_ value: String, x: CGFloat, top: CGFloat, width: CGFloat, height: CGFloat = 16,
          size: CGFloat = 8, bold: Bool = false, right: Bool = false, red: Bool = false) {
    let paragraph = NSMutableParagraphStyle()
    paragraph.alignment = right ? .right : .left
    paragraph.lineBreakMode = .byWordWrapping
    let attributes: [NSAttributedString.Key: Any] = [
        .font: bold ? NSFont.boldSystemFont(ofSize: size) : NSFont.systemFont(ofSize: size),
        .foregroundColor: red ? NSColor.systemRed : NSColor.black,
        .paragraphStyle: paragraph
    ]
    (value as NSString).draw(in: CGRect(x: x, y: top, width: width, height: height), withAttributes: attributes)
}

func beginPage() {
    context.beginPDFPage(nil)
    context.saveGState()
    context.translateBy(x: 0, y: page.height)
    context.scaleBy(x: 1, y: -1)
    NSGraphicsContext.current = NSGraphicsContext(cgContext: context, flipped: true)
    pageNumber += 1
    text("Balansrapport", x: 24, top: 22, width: 450, height: 25, size: 18, bold: true)
    text(report.metadata.FNAMN + " · " + report.metadata.ORGNR, x: 24, top: 51, width: 540, size: 10)
    text("Räkenskapsår: \(report.metadata.start)–\(report.metadata.end) · Senaste verifikation: \(report.metadata.latest_voucher)", x: 24, top: 69, width: 550)
    text("Källa: sie4-export.se · Genererad: \(report.metadata.generated.prefix(10)) · Belopp i SEK", x: 24, top: 84, width: 550)
    y = 108
    let headings = ["Konto", "Kontonamn", "Ingående balans", "Period", "Utgående balans"]
    for i in 0..<5 { text(headings[i], x: columns[i], top: y, width: widths[i], size: 7.4, bold: true, right: i >= 2) }
    y += 19
}

func endPage() {
    text("Sida \(pageNumber) · SIE SHA256: \(report.metadata.sha256.prefix(20))…", x: 24, top: 816, width: 550, size: 7)
    context.restoreGState()
    context.endPDFPage()
}

beginPage()
for row in report.rows {
    let isAccount = row.kind == "account"
    let isCheck = row.kind == "check"
    let rowHeight: CGFloat = isCheck ? 26 : isAccount && row.label.count > 59 ? 23 : 14
    if y + rowHeight > 770 { endPage(); beginPage() }
    if row.kind == "major" || row.kind == "section" {
        context.setFillColor(NSColor(calibratedWhite: 0.94, alpha: 1).cgColor)
        context.fill(CGRect(x: 22, y: y - 1, width: 552, height: rowHeight))
    }
    text(row.account, x: columns[0], top: y, width: widths[0])
    text(row.label, x: columns[1], top: y, width: widths[1], height: rowHeight,
         size: isCheck ? 7.4 : 8, bold: !isAccount, red: isCheck)
    if let values = row.values {
        for (i, value) in values.enumerated() {
            let number = NSDecimalNumber(string: value)
            text(formatter.string(from: number)!, x: columns[i + 2], top: y, width: widths[i + 2],
                 bold: !isAccount, right: true, red: isCheck && number != NSDecimalNumber.zero)
        }
    }
    y += rowHeight
}
y += 14
for note in report.notes {
    if y + 35 > 795 { endPage(); beginPage() }
    text(note, x: 24, top: y, width: 547, height: 35, size: 7.5)
    y += 35
}
endPage()
context.closePDF()
