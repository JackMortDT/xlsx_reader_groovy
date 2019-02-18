@Grab(group='org.apache.poi', module='poi', version='3.8')
@Grab(group='org.apache.poi', module='poi-ooxml', version='3.8')

import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.util.*
import org.apache.poi.ss.usermodel.*
import java.io.*

class GroovyExcelParser {
  //http://poi.apache.org/spreadsheet/quick-guide.html#Iterator

  def parse(path) {
      InputStream inp = new FileInputStream(path)
      Workbook wb = WorkbookFactory.create(inp);
      Sheet sheet = wb.getSheetAt(0);

      Iterator<Row> rowIt = sheet.rowIterator()
      Row row = rowIt.next()
      def headers = getRowData(row)

      def rows = []
      while(rowIt.hasNext()) {
            row = rowIt.next()
            rows << getRowData(row)
          }
      [headers, rows]
    }

  def getRowData(Row row) {
      def data = []
      for (Cell cell : row) {
            getValue(row, cell, data)
          }
      data
    }

  def getValue(Row row, Cell cell, List data) {
      def rowIndex = row.getRowNum()
      def colIndex = cell.getColumnIndex()
      def value = ""
      switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
              value = cell.getRichStringCellValue().getString();
              break;
            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(cell)) {
                          value = cell.getDateCellValue();
                      } else {
                                  value = cell.getNumericCellValue();
                              }
              break;
            case Cell.CELL_TYPE_BOOLEAN:
              value = cell.getBooleanCellValue();
              break;
            case Cell.CELL_TYPE_FORMULA:
              value = cell.getCellFormula();
              break;
            default:
              value = ""
          }
      data[colIndex] = value
      data
    }

  public static void main(String[]args) {
      def filename = '/Users/luissas/Downloads/CambiodeCalificaciónde2019a2018RutasDocentesV1.xlsx'
      GroovyExcelParser parser = new GroovyExcelParser()
      def (headers, rows) = parser.parse(filename)
      def file = new File('/Users/luissas/test.txt')

      rows.each{
        file << "update \"APPDOCENTE\".\"ASSIGNEDEVALUATION\" SET starttime = TO_TIMESTAMP('31-12-2018','DD-MM-YYYY'), lastupdated = TO_TIMESTAMP('31-12-2018','DD-MM-YYYY') where id = (select id from assignedevaluation where evaluation_id = 1 and userrd_id = (select id from userrd where username = '${it[5]}')); \n"
      }
      file.createNewFile()
      println "🤯"
    }
}
