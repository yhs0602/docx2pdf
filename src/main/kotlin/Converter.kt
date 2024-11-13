import com.lowagie.text.Font
import com.lowagie.text.pdf.BaseFont
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.awt.Color
import java.io.File
import java.io.OutputStream


fun docx2pdf(inputFile: File) {
    // Replace the extension of the file
    val filenameWithoutExtension = inputFile.nameWithoutExtension
    val outputFile = File(inputFile.parent, "$filenameWithoutExtension.pdf")
    val koreanFont = BaseFont.createFont(
        "/NanumGothic.ttf", BaseFont.IDENTITY_H,
        BaseFont.EMBEDDED
    )
    val symbolFont = BaseFont.createFont(
        "/MaterialIcons-Regular.ttf", BaseFont.IDENTITY_H,
        BaseFont.EMBEDDED
    )
    val boldKoreanFont = BaseFont.createFont(
        "/NanumGothicBold.ttf", BaseFont.IDENTITY_H,
        BaseFont.EMBEDDED
    )
    try {
        val fis = inputFile.inputStream()
        // Read docx using Apache POI
        val docx = XWPFDocument(fis)
        // Convert to pdf
        val pdfOptions = PdfOptions.create().fontEncoding("UTF-8")
            .fontProvider { familyName: String?, encoding: String?, size: Float, style: Int, color: Color? ->
                val fontToUse: BaseFont =
                    when (familyName) {
                        "Wingdings" -> symbolFont
                        else -> {
                            println("Requested font family: $familyName, encoding: $encoding style: $style")
                            if ((style > 0) && (style and 1) == 1) {
                                koreanFont
                            } else {
                                println("Requested font family: $familyName, encoding: $encoding")
                                koreanFont
                            }
                        }
                    }
                val font = Font(fontToUse, size, style, color)
                if (familyName != null) {
                    font.setFamily(familyName)
                }
                font
            }
        val out: OutputStream = outputFile.outputStream()
        PdfConverter.getInstance().convert(docx, out, pdfOptions)
    } catch (e: Exception) {
        e.printStackTrace()
    }
}