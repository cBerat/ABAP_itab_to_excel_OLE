*&---------------------------------------------------------------------*
*& Report ZBC_EXCEL_OLE
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT ZBC_EXCEL_OLE.

"bu program calýsýnca ilgili veri tabaný tablosu doldurulur
"excel programý acýlýp ici db verileriyle doldurulur

DATA: application type ole2_object,
      workbook    type ole2_object,
      sheet       type ole2_object,
      cells       type ole2_object,
      gt_scarr    type table of scarr,
      gs_scarr    type scarr,
      gv_row      type i.

START-OF-SELECTION.

  " workbook olusturma islemi
  create OBJECT application 'excel.application'.
  set PROPERTY OF application 'visible' = 1.
  call method of application 'Workbooks' = workbook.
  call method of workbook 'Add'.
  " sheet olusturma islemi
  call METHOD of application 'Worksheets' = sheet
  exporting #1 = 1.
  call method of sheet 'Activate' .
  set PROPERTY OF sheet 'Name' = 'Sheet1'.


" burada baslýklar belirlendi fill cell formunda 1. satýr 1. sutuna üst birim yollandý
  perform fill_cell USING 1 1 'Ust Birim'.
  perform fill_cell using 1 2 'Kýsa Taným'."1. satýr 2. sutuna kýsa taným yollandý
  perform fill_cell using 1 3 'Havayolu sirketinin adi'. "1. satýr 3. sutuna Havayolu sirketinin adi yollandý
  perform fill_cell using 1 4 'PB'.
  perform fill_cell using 1 5 'URL'.

  select * from scarr into table gt_scarr.

  LOOP AT gt_scarr into gs_scarr.
    gv_row = sy-TABIX + 1. " her satýrý dolassýn diye yapýldý.
    perform fill_cell using gv_row 1 gs_scarr-mandt. " satýr satýr degerler excele islenmis oldu.
    perform fill_cell using gv_row 2 gs_scarr-CARRID.
    perform fill_cell using gv_row 3 gs_scarr-CARRNAME.
    perform fill_cell using gv_row 4 gs_scarr-CURRCODE.
    perform fill_cell using gv_row 5 gs_scarr-URL.
  ENDLOOP.

form fill_cell using row col val.
  call method of sheet 'Cells' = cells EXPORTING #1 = row #2 = col.
  set PROPERTY OF cells 'Value' = val.
endform.
