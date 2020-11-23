package test

import org.apache.poi.ss.usermodel.{Cell, DataFormatter, Row, Sheet, WorkbookFactory,
  HorizontalAlignment, BorderStyle, FillPatternType, CellType, VerticalAlignment}

import scala.collection.mutable.ListBuffer
import scala.collection.mutable.{Map, MultiMap, HashMap}
import java.io._
import play.api.libs.json._
import java.util.zip._



import scala.collection.mutable

object Main extends App {

  def getAsJson(sheet: Sheet, path: String): JsValue = {
    val mode = "excel"
    val max_rows = 5000
    val max_cols = 1000
    val (directory, fname) = getDirFname(path)
    val repeatMergedCells = true
    val fmt = new DataFormatter();
//    println(s"${sheet.getSheetName()}, ${sheet.getPhysicalNumberOfRows()}, ${sheet.getNumMergedRegions()}")
    var mergedRegions = new ListBuffer[String]()
    for (i <- 0 to sheet.getNumMergedRegions()-1){
      var mergedRegion = sheet.getMergedRegion(i)
      mergedRegions += mergedRegion.formatAsString()
    }
    val mergedRegionsInfo = parseMergedCells(mergedRegions)
    val mergedRegionsMap = mergedRegionsInfo._1
    val mergedRegionsSize = mergedRegionsInfo._2
    var table_arr = new ListBuffer[ListBuffer[String]]()
    var type_arr = new ListBuffer[ListBuffer[String]]()
    var feature_arr = new ListBuffer[ListBuffer[List[Double]]]()


    val N = sheet.getPhysicalNumberOfRows()
    var M = 0
    println(sheet.getSheetName(), N)
    for (i <- 0 to math.min(N, max_rows)-1) {
      val row = sheet.getRow(i)
      var row_vals = new ListBuffer[String]()
      var row_types = new ListBuffer[String]()
      var row_features = new ListBuffer[List[Double]]()
      if (row != null) {
        val ncol = math.min(row.getLastCellNum(), max_cols)
        M = scala.math.max(M, ncol)
        for (j <- 0 to ncol - 1) {
          val cell = row.getCell(j)
          if (cell != null) {
            var value = fmt.formatCellValue(cell)
            val type_ = cell.getCellType()
            if (type_ == CellType.FORMULA) {
              cell.getCachedFormulaResultType() match{
                case CellType.BOOLEAN => value = cell.getBooleanCellValue().toString
                case CellType.NUMERIC => value = cell.getNumericCellValue().toString
                case CellType.STRING => value = cell.getRichStringCellValue().toString
                case _ => "the prev value is good to go"
              }
            }
            row_vals += value
            row_types += type_.toString()
            val cell_features = extractCellFeatures(cell, value, i, j, mergedRegionsMap, mergedRegionsSize)
            row_features += getFeatureVector(cell_features, mode)
          }else {
            row_vals += ""
            row_types += ""
          }
        }
      }
      table_arr += row_vals
      type_arr += row_types
      feature_arr += row_features
    }
    for (row <- table_arr) { // fill in the incomplete rows
      if (row.length < M) {
        for (i <- row.length to M-1) {
          row += ""
        }
      }
    }
    for (row <- feature_arr) { // fill in the incomplete rows
      if (row.length < M) {
        for (i <- row.length to M-1) {
          row += getNullFeatureVector(mode)
        }
      }
    }
    if (repeatMergedCells) {
      table_arr = adjustMergedBlocksTarr(table_arr, mergedRegionsMap)
      feature_arr = adjustMergedBlocksFeat(feature_arr, mergedRegionsMap)
    }
    return createJson(directory, fname, sheet.getSheetName(),
      table_arr, feature_arr, N, M, mergedRegions,
      if(mode == "excel") getFeatureNames() else getFeatureNamesReduced())
  }

  def adjustMergedBlocksTarr(tarr: ListBuffer[ListBuffer[String]],
                         merged_region_map: Map[(Int, Int), (Int, Int)]): ListBuffer[ListBuffer[String]] = {
    for (i <- 0 to tarr.length -1){
      for (j <- 0 to tarr(i).length - 1) {
        if (merged_region_map.contains((i,j))){
          val orig_coord = merged_region_map.getOrElse((i,j), (i,j))
          tarr(i)(j) = tarr(orig_coord._1)(orig_coord._2)
        }
      }
    }
    return tarr
  }

  def adjustMergedBlocksFeat(feat_arr: ListBuffer[ListBuffer[List[Double]]],
                         merged_region_map: Map[(Int, Int), (Int, Int)]): ListBuffer[ListBuffer[List[Double]]] = {
    for (i <- 0 to feat_arr.length -1){
      for (j <- 0 to feat_arr(i).length - 1) {
        if (merged_region_map.contains((i,j))){
          val orig_coord = merged_region_map.getOrElse((i,j), (i,j))
          feat_arr(i)(j) = feat_arr(orig_coord._1)(orig_coord._2)
        }
      }
    }
    return feat_arr
  }

  def processFile(file_name: String): List[JsValue] = {
    val annotation_sheets = Set("Range_Annotations_Data", "Annotation_Status_Data")
    val f = new File(file_name)
    val workbook = WorkbookFactory.create(f, "", true)
    var res = new ListBuffer[JsValue]
    for (i <- 0 to workbook.getNumberOfSheets()-1) {
      val sheet = workbook.getSheetAt(i) // Assuming they're in the first sheet here.
      res += getAsJson(sheet, file_name)
    }
    if (res.map(_("table_id")).map(_.toString()).toList.exists(x => {x.contains("Range_Annotations_Data")})) {
      res = addAnnotations(res)
    }
//    else {
//      res.map(_("file_name")).map(_.toString()).foreach(print)
//    }
    workbook.close()
    return res.toList
  }

  def addAnnotations(jobjects: ListBuffer[JsValue]): ListBuffer[JsValue] = {
    var res = new ListBuffer[JsValue]()
    val index = jobjects.map(_("table_id")).map(_.toString().contains("Range_Annotations_Data")).indexOf(true)
    val ann_sheet = jobjects(index)
//    println(ann_sheet("table_array"))
    val ann_map = parseAnnotations(ann_sheet("table_array").as[JsArray], ann_sheet("num_rows").as[Int])
    for (jobj <- jobjects) {
      var jobj_new = jobj
      val sname = jobj("table_id").as[String]
      val n = jobj("num_rows").as[Int]
      val m = jobj("num_cols").as[Int]
      if (!sname.contains("Range_Annotations_Data") && !sname.contains("Annotation_Status_Data") && n > 0 && m > 0) {
        if (ann_map.contains(sname)) {
          // initialize annotations
          var annotations = Array.ofDim[String](n, m)
          for ((begin, end, label) <- ann_map.getOrElse(sname, null)) {
            for (i <- begin._1 to math.min(end._1,n-1)){
              for (j <- begin._2 to math.min(end._2,m-1)) {
                annotations(i)(j) = label
              }
            }
          }
          jobj_new = jobj.as[JsObject] + ("annotations" -> Json.toJson(annotations))
        }
        res += jobj_new
      }
    }
    return res
  }

  def parseAnnotations(ann: JsArray, n: Int): Map[String, ListBuffer[((Int, Int), (Int, Int), String)]] = {
    var res: Map[String, ListBuffer[((Int, Int), (Int, Int), String)]] = Map()
    for (i <- 1 to n-1) {
      val row = ann(i)
      val (sname, begin, end, label) = parseAnnotation(row.as[List[String]])
      if (label != "table") {
        if (res.contains(sname)) {
          var x = res.getOrElse(sname, null)
          x += ((begin, end, label))
        } else {
          res += (sname -> ListBuffer((begin, end, label)))
        }
      }
    }
    return res
  }

  def parseAnnotation(row: List[String]): (String, (Int, Int), (Int, Int), String) = {
    // name: 0, label: 2, range:4
    val sname = row(0)
    val label = row(2).toLowerCase()
    val (begin, end) = parseAnnotationRange(row(4))
    return (sname, begin, end, label)
  }

  def parseAnnotationRange(range: String): ((Int, Int), (Int, Int)) = {
    if (range.contains(':')) {
      val args = range.split(':')
      val begin = toIJ(args(0).filterNot(_=='$'))
      val end = toIJ(args(1).filterNot(_=='$'))
      return (begin, end)
    } else {
      val begin = toIJ(range.filterNot(_=='$'))
      val end = begin
      return (begin, end)
    }
  }

  def getDirFname(path: String): (String, String) = {
    val res = path.splitAt(path.lastIndexOf('/'))
    return res
  }

  def createJson(directory: String,
                 fname: String,
                 sheet_name: String,
                 tarr: ListBuffer[ListBuffer[String]],
                 feature_arr: ListBuffer[ListBuffer[List[Double]]],
                 num_rows: Int,
                 num_cols: Int,
                 merged_regions: ListBuffer[String],
                 feature_names: Seq[String]): JsValue = {
    val json: JsValue = Json.obj (
      "table_array" -> Json.toJson(tarr),
      "feature_array" -> Json.toJson(feature_arr),
      "num_rows" -> num_rows,
      "num_cols" -> num_cols,
      "merged_regions" -> Json.toJson(merged_regions),
      "directory" -> directory,
      "table_id" -> sheet_name,
      "file_name" -> fname,
      "feature_names" -> Json.toJson(feature_names)
    )
    // println(Json.stringify(json))
    // println(Json.prettyPrint(json))
    return json
  }

  def getFeatureNamesReduced(): Seq[String] = {
    return Seq("all_upper", "capitalized",
      "contains_colon", "first_char_num",
      "first_char_special", "first_col_num", "first_row_num",
      "in_year_range",
      "is_alpha",
      "leading_spaces", "length",
      "num_of_neighbors=0", "num_of_neighbors=1",
      "num_of_neighbors=2", "num_of_neighbors=3", "num_of_neighbors=4",
      "punctuations",
      "special_chars", "words",
      "words_like_table", "words_like_total")
  }

  def getFeatureNames(): Seq[String] = {
    return Seq("all_upper", "border_bottom_type=0", "border_left_type=0",
      "border_right_type=0", "border_right_type=2", "border_top_type=0",
      "border_top_type=1", "capitalized", "cell_borders=0",
      "cell_type=0", "cell_type=1", "cell_type=2", "contains_colon",
      "fill_patern=0", "first_char_num",
      "first_char_special", "first_col_num", "first_row_num", "font_height",
      "h_alignment=0", "h_alignment=2", "in_year_range",
      "indentation", "is_aggr_formula=1", "is_alpha", "is_bold",
      "is_font_color_default", "is_wraptext", "leading_spaces", "length",
      "num_of_cells", "num_of_neighbors=0", "num_of_neighbors=1",
      "num_of_neighbors=2", "num_of_neighbors=3", "num_of_neighbors=4",
      "punctuations", "is_italic",
      "special_chars", "underline_type=0", "v_alignment=2", "words",
      "words_like_table", "words_like_total")
  }

  def toIJ(s: String): (Int, Int) = {
    val alphabet = ('a' to 'z').toList.map(_.toString)
    val double_alphabet = alphabet.flatMap(x => alphabet.map(y => (x, y))).map{case (a,b) => s"${a}${b}"}
    val triple_alphabet = double_alphabet.flatMap(x => alphabet.map(y => (x, y))).map{case (a,b) => s"${a}${b}"}
    val colNames = (
        alphabet ++
        double_alphabet ++
        triple_alphabet
      )
    val pattern = "([a-z]+)([0-9]+)".r
    val pattern(colstring, rowstring) = s.toLowerCase()
    val row = rowstring.toInt - 1
    val col = colNames.indexOf(colstring)
    return (row, col)
  }

  def parseMergedCells(mergedCells: ListBuffer[String]): (Map[(Int, Int), (Int, Int)], Map[(Int, Int), Int]) = {
    var mergedCellMap:Map[(Int, Int), (Int, Int)] = Map()
    var mergedSizeMap:Map[(Int, Int), Int] = Map()
    for (block <- mergedCells){
      var size = 0
      val begin_end = block.split(":").map(toIJ)
      val orig_i = begin_end(0)._1
      val orig_j = begin_end(0)._2
      for (i <- begin_end(0)._1 to begin_end(1)._1){
        for (j <- begin_end(0)._2 to begin_end(1)._2){
          mergedCellMap += ((i,j) -> (orig_i, orig_j))
          size += 1
        }
      }
      mergedSizeMap += ((orig_i, orig_j) -> size)
    }
    return (mergedCellMap, mergedSizeMap)
  }

  def getFeatureVector(features:Map[String, Double], mode:String): List[Double] = {
    mode match {
      case "excel" => return getFeatureNames().map(features.getOrElse(_, 0.0)).toList
      case "csv" => return getFeatureNamesReduced().map(features.getOrElse(_, 0.0)).toList
    }
  }

  def getNullFeatureVector(mode:String): List[Double] = {
    mode match {
      case "excel" => return getFeatureNames().map(x => 0.0).toList
      case "csv" => return getFeatureNamesReduced().map(x => 0.0).toList
    }
  }

  def extractCellFeatures(
                           cell: Cell,
                           value: String,
                           i: Int,
                           j: Int,
                           mergedRegionsMap: Map[(Int, Int), (Int, Int)],
                           mergedRegionsSize: Map[(Int, Int), Int]
                         ): Map[String, Double] = {
    var res:Map[String, Double] = Map()
    /* Syntactic Features*/
    res += ("all_upper" -> (if(isAllUpper(value)) 1.0 else 0))
    // contains_colon
    res += ("contains_colon" -> (if(hasColon(value)) 1.0 else 0))
    // capitalized
    res += ("capitalized" -> (if(hasUpper(value)) 1.0 else 0))
    // first_char_num
    res += ("first_char_num" -> (if(isFirstCharNum(value)) 1.0 else 0))
    // first_char_special
    res += ("first_char_special" -> (if(isFirstCharSpecial(value)) 1.0 else 0))
    // is_alpha
    res += ("is_alpha" -> (if(isAlpha(value)) 1.0 else 0))
    // leading_spaces
    res += ("leading_spaces" -> countLeadingSpaces(value))
    // length
    res += ("length" -> getLength(value))
    // punctuations
    res += ("punctuations" -> (if(hasPunctuations(value)) 1.0 else 0))
    // special_chars
    res += ("special_chars" -> (if(hasSpecialChar(value)) 1.0 else 0))
    // words
    res += ("words" -> countWords(value))
    // words_like_table
    res += ("words_like_table" -> (if(hasWordsLikeTable(value)) 1.0 else 0))
    // words_like_total
    res += ("words_like_total" -> (if(hasWordsLikeTotal(value)) 1.0 else 0))
    // in_year_range
    res += ("in_year_range" -> (if(isInYearRange(value)) 1.0 else 0))

    val orig_loc = getFirstRowCol(i, j, mergedRegionsMap)
    // first_col_num
    res += ("first_col_num" -> orig_loc._2)
    // first_row_num
    res += ("first_row_num" -> orig_loc._1)
    // num_of_cells
    res += ("num_of_cells" -> getRegionSize(i, j, mergedRegionsMap, mergedRegionsSize))

    val sheet = cell.getSheet()
    val workbook = sheet.getWorkbook()
    val cell_style = cell.getCellStyle()
    val cell_type = cell.getCellType()

    val border_top = cell_style.getBorderTop()
    val border_btm = cell_style.getBorderBottom()
    val border_lft = cell_style.getBorderLeft()
    val border_rght = cell_style.getBorderRight()
    val type0_borders = Set(BorderStyle.NONE)
    val type1_borders = Set(BorderStyle.THIN)

    val hor_alg = cell_style.getAlignment()
    val vert_alg = cell_style.getVerticalAlignment()
    val type0_align = Set(HorizontalAlignment.GENERAL, HorizontalAlignment.LEFT,
      HorizontalAlignment.JUSTIFY)
    val type2_align = Set(HorizontalAlignment.CENTER, HorizontalAlignment.CENTER_SELECTION,
      HorizontalAlignment.FILL, HorizontalAlignment.DISTRIBUTED)

    val type2_valign = Set(VerticalAlignment.CENTER)

    val type0_cell = Set(CellType._NONE, CellType.BLANK)
    val type1_cell = Set(CellType.BOOLEAN, CellType.STRING)
    val type2_cell = Set(CellType.NUMERIC, CellType.FORMULA, CellType.ERROR)
    val type_formula_cells = Set(CellType.FORMULA, CellType.ERROR)

    val font = workbook.getFontAt(cell_style.getFontIndexAsInt())

    def my_get_cell(ii:Int, jj:Int): Cell ={
      if (ii < 0 || jj < 0 ) return null
      return if (sheet.getRow(ii) != null) sheet.getRow(ii).getCell(jj) else null
    }
    val neigbors = List(
      my_get_cell(i-1, j),
      my_get_cell(i, j-1),
      my_get_cell(i+1, j),
      my_get_cell(i, j+1)
    )
    val num_neghbors = neigbors.map(c => if(c != null) 1 else 0).sum
//    println(neigbors, num_neghbors, value)

//    println(font.getUnderline(), cell_style.getFillBackgroundColor())

    /* Stylistic Features */
    // "border_bottom_type=0",
    res += ("border_bottom_type=0" -> (if(type0_borders.contains(border_btm)) 1.0 else 0))
    // "border_left_type=0",
    res += ("border_left_type=0" -> (if(type0_borders.contains(border_lft)) 1.0 else 0))
    // "border_right_type=0",
    res += ("border_right_type=0" -> (if(type0_borders.contains(border_rght)) 1.0 else 0))
    // "border_right_type=2",
    res += ("border_right_type=2" -> (if(!type0_borders.contains(border_rght) && !type1_borders.contains(border_rght)) 1.0 else 0))
    // "border_top_type=0",
    res += ("border_top_type=0" -> (if(type0_borders.contains(border_top)) 1.0 else 0))
    // "border_top_type=1",
    res += ("border_top_type=1" -> (if(type1_borders.contains(border_top)) 1.0 else 0))
    // "cell_borders=0",
    res += ("cell_borders=0" -> (if(List(border_btm, border_lft, border_rght, border_top).forall(type0_borders.contains(_))) 1.0 else 0))
    // "cell_type=0",
    res += ("cell_type=0" -> (if(type0_cell.contains(cell_type)) 1.0 else 0))
    // "cell_type=1",
    res += ("cell_type=1" -> (if(type1_cell.contains(cell_type)) 1.0 else 0))
    // "cell_type=2",
    res += ("cell_type=2" -> (if(type2_cell.contains(cell_type)) 1.0 else 0))
    // "fill_patern=0",
    res += ("fill_patern=0" -> (if(cell_style.getFillPattern() == FillPatternType.NO_FILL) 1.0 else 0))
    // "font_height",
    res += ("font_height" -> font.getFontHeightInPoints())
    // "h_alignment=0",
    res += ("h_alignment=0" -> (if(type0_align.contains(hor_alg)) 1.0 else 0))
    // "h_alignment=2",
    res += ("h_alignment=2" -> (if(type2_align.contains(hor_alg)) 1.0 else 0))
    // "indentation",
    res += ("indentation" -> cell_style.getIndention())
    // "is_aggr_formula=1", TODO: check whether is for special formula
    res += ("is_aggr_formula=1" -> (if(type_formula_cells.contains(cell_type)) 1.0 else 0))
    // "is_bold",
    res += ("is_bold" -> (if(font.getBold()) 1.0 else 0))
    // "is_italic",
    res += ("is_italic" -> (if(font.getItalic()) 1.0 else 0))
    // "is_font_color_default",
    res += ("is_font_color_default" -> (if(font.getColor() == 0) 1.0 else 0))
    // "is_wraptext",
    res += ("is_wraptext" -> (if(cell_style.getWrapText()) 1.0 else 0))
    // "num_of_neighbors=0",
    res += ("num_of_neighbors=0" -> (if(num_neghbors == 0) 1.0 else 0))
    // "num_of_neighbors=1",
    res += ("num_of_neighbors=1" -> (if(num_neghbors == 0) 1.0 else 0))
    // "num_of_neighbors=2",
    res += ("num_of_neighbors=2" -> (if(num_neghbors == 0) 1.0 else 0))
    // "num_of_neighbors=3",
    res += ("num_of_neighbors=3" -> (if(num_neghbors == 0) 1.0 else 0))
    // "num_of_neighbors=4",
    res += ("num_of_neighbors=4" -> (if(num_neghbors == 0) 1.0 else 0))
    // "underline_type=0",
    res += ("underline_type=0" -> (if(font.getUnderline() != 0) 1.0 else 0))
    // "v_alignment=2",
    res += ("v_alignment=2" -> (if(type2_valign.contains(vert_alg)) 1.0 else 0))
    // "ref_val_type=0", TODO: was not even in python code
    return res
  }

  def isAllUpper(value: String): Boolean = {
    // val ordinary=(('A' to 'Z') ++ ('0' to '9')).toSet
    return value.matches("^[A-Z\\s]*$")
  }

  def hasUpper(value: String): Boolean = {
    return value.exists(_.isUpper)
  }

  def hasColon(value: String): Boolean = {
    return value.exists(_ == ':')
  }

  def isAlpha(value: String): Boolean = {
    return value.forall(_.isLetterOrDigit)
  }

  def isFirstCharNum(value: String): Boolean = {
    if (value.length > 0) {
      val numbers = (0 to 9).toSet
      return numbers.contains(value(0).toInt)
    } else {
      return false
    }
  }

  def isFirstCharSpecial(value: String): Boolean = {
    if (value.length > 0) {
      val special_chars = (0 to 9).toSet
      return ! value(0).toString.matches("[a-zA-Z0-9\\s]")
    } else {
      return false
    }
  }

  def isInYearRange(value: String): Boolean = {
    try {
      val int_val = value.toInt
      if ((value.toFloat - int_val) < 0.001 && int_val > 1800 && int_val < 2091) {
        return true
      }
      return false
    } catch {
      case e: Exception => return false
    }
  }

  def countLeadingSpaces(value: String): Int = {
    val res = value.indexOf(value.trim())
    return res
  }

  def getLength(value: String): Int = {
    return value.trim().length
  }

  def hasPunctuations(value: String): Boolean = {
    val puncs = "!\"#$%&\\'()*+,-./:;<=>?@[\\\\]^_`{|}~".toCharArray.toSet
    return value.toCharArray.exists(puncs.contains(_))
  }

  def hasSpecialChar(value: String): Boolean = {
    val pattern = """[^\x20-\x7E]+""".r
    val res = pattern.findFirstMatchIn(value) match {
      case Some(_) => true
      case None => false
    }
    return res
  }

  def countWords(value: String): Int = {
    val res = value.trim().split("\\s+").length
    return res
  }

  def hasWordsLikeTable(value: String): Boolean = {
    val wordsLikeTable = Seq("table").toSet
    val res = value.trim().split("\\s+").exists(wordsLikeTable.contains(_))
    return res
  }

  def hasWordsLikeTotal(value: String): Boolean = {
    val totalLike = Seq("total", "average", "maximum", "minimum").toSet
    val res = value.trim().split("\\s+").exists(totalLike.contains(_))
    return res
  }

  def getFirstRowCol(i: Int, j:Int, mergedRegionsMap: Map[(Int, Int), (Int, Int)]): (Int, Int) = {
    val res = mergedRegionsMap.getOrElse((i,j), (i,j))
    return res
  }

  def getRegionSize(i: Int, j:Int,
                    mergedRegionsMap: Map[(Int, Int), (Int, Int)],
                    mergedRegionsSize: Map[(Int, Int), Int]): Int = {
    val coord = mergedRegionsMap.getOrElse((i,j), (i,j))
    val res = mergedRegionsSize.getOrElse(coord, 1)
    return res
  }


  def recursiveListFiles(f: File): Array[File] = {
    var these = f.listFiles
    these ++= these.filter(_.isDirectory).flatMap(recursiveListFiles)
    return these
  }

  def listFiles(path: String): Array[String] = {
    val f = new File(path)
    val files = recursiveListFiles(f)
    files.map(_.getAbsolutePath()).filter(_.matches(".*xlsx?$"))
  }

  def createDataset(entry_path: String, out_path: String): Unit = {
    val files = listFiles(entry_path)
    var fos = new FileOutputStream(out_path)
    var gzos = new GZIPOutputStream( fos )
    var writer = new PrintWriter(gzos)
    var counter:Int = 0
    val offset = 0
    for (fpath <- files.slice(offset, files.length-1)) {
      counter += 1
      println(s"processing ${counter+offset}/${files.length}  ${fpath}")
      try {
        for (jobj <- processFile(fpath)) {
          writer.write(Json.stringify(jobj) + "\n")
        }
      } catch {
        case e: Exception => println("failed")
      }
    }
    writer.close()
    gzos.close()
  }

  createDataset("/media/majid/data/data_new/fbi_tables_all/",
    "/media/majid/data/Download/fbi_tables_all.jl.gz")
}


