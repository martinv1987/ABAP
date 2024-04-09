FUNCTION ZSD_DESCARGA_XLSX.
*"----------------------------------------------------------------------
*"*"Interfase local
*"  IMPORTING
*"     VALUE(I_LIBROS) TYPE  ZTTY_LIBROS_XLSX
*"     VALUE(I_DATA) TYPE  ZTTY_DATA_XLSX
*"     VALUE(I_MERGECELLS) TYPE  ZTTY_MERGECELLS_XLSX OPTIONAL
*"     VALUE(I_TOTALES) TYPE  ZTTY_TOTALES_XLSX OPTIONAL
*"     VALUE(I_CELLCOLOR) TYPE  ZEXCEL_STYLE_COLOR_ARGB OPTIONAL
*"----------------------------------------------------------------------

DATA:
      lo_excel           TYPE REF TO zcl_excel,
      lo_worksheet       TYPE REF TO zcl_excel_worksheet,
      lo_style_bold      TYPE REF TO zcl_excel_style,
      lo_style_blue      TYPE REF TO zcl_excel_style,
      lo_style_center    TYPE REF TO zcl_excel_style,
      lo_style_blue_dark TYPE REF TO zcl_excel_style,
      lo_style_underline TYPE REF TO zcl_excel_style,
      lo_style_filled    TYPE REF TO zcl_excel_style,
      lo_style_border    TYPE REF TO zcl_excel_style,
      lo_style_button    TYPE REF TO zcl_excel_style,
      lo_border_dark     TYPE REF TO zcl_excel_style_border,
      lo_border_light    TYPE REF TO zcl_excel_style_border,
      lo_column    TYPE REF TO zcl_excel_column,
      lt_rawdata TYPE solix_tab,
      vl_data TYPE xstring,
      vl_texto TYPE char255,
      vl_row TYPE i,
      vl_counter TYPE i,
      lv_bin_size TYPE i,
      buffer_zip TYPE xstring,
      vl_filename TYPE string,
      vl_path TYPE string,
      vl_fullpath TYPE string,
      lv_style_bold_guid             TYPE zexcel_cell_style,
      lv_style_center_guid           TYPE zexcel_cell_style,
      lv_style_blue_guid             TYPE zexcel_cell_style,
      lv_style_blue_guid_dark        TYPE zexcel_cell_style,
      lv_style_underline_guid        TYPE zexcel_cell_style,
      lv_style_filled_guid           TYPE zexcel_cell_style,
      lv_style_filled_green_guid     TYPE zexcel_cell_style,
      lv_style_border_guid           TYPE zexcel_cell_style,
      lv_bytecount TYPE i,
      lt_file_tab  TYPE solix_tab,
      lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c,
      lo_row TYPE REF TO zcl_excel_row,
      t_data_tab TYPE TABLE OF x255,
      wa_input TYPE ZSTSD_R0150_XLS_MAIL,
      lex_ecx_excel TYPE REF TO ZCX_EXCEL,
      vl_bin_img TYPE xstring,
      vl_letra TYPE string,
      vl_value TYPE zexcel_cell_value,
      vl_cantidad_reg TYPE i,
      vl_current_row TYPE i,
      vl_letter TYPE string,
      vl_number TYPE i,
      vl_loopletra TYPE i,
      vl_cellcolor TYPE ZEXCEL_STYLE_COLOR_ARGB,
      lo_excel_writer         TYPE REF TO zif_excel_writer,
      vl_valor_fs type string,
      i_data_aux TYPE STANDARD TABLE OF ZSTSD_DATA_XLSX,
      vl_saltodelinea TYPE flag,
      go_table type ref to cl_abap_tabledescr,
      go_struct type ref to cl_abap_structdescr,
      gt_comp   type abap_component_tab,
      gs_comp   type abap_componentdescr.


* Se asigna un valor por defecto a las celdas que sirven de cabecera
if i_cellcolor is initial.
   vl_cellcolor = '0033BBFF'.
else.
   vl_cellcolor = i_cellcolor.
endif.

FIELD-SYMBOLS: <fs_valor> TYPE any.

*Se abre el cuadro de diálogo para la búsqueda de archivo
  CALL METHOD cl_gui_frontend_services=>file_save_dialog
      EXPORTING
          window_title = 'Escoja una ruta de acceso'
          default_file_name = 'archivo.xlsx'
          default_extension = '*.XLSX'
          file_filter = 'Archivo XLSX (*.*)|*.*|Text files (*.xlsx)|*.xlsx'
      CHANGING
          filename = vl_filename
          path = vl_path
          fullpath = vl_fullpath
      EXCEPTIONS
          cntl_error = 1
          error_no_gui = 2
          not_supported_by_gui = 3
      OTHERS = 4.

IF vl_fullpath is initial.

*    raise EX_CANCEL.

ENDIF.

  try.
      CREATE OBJECT lo_excel."<Se crea el objeto

      CREATE OBJECT lo_border_dark."Se crea borde y se asigna su tamaño y color
      lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
      lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.

      CREATE OBJECT lo_border_light.
      lo_border_light->border_color-rgb = zcl_excel_style_color=>c_gray.
      lo_border_light->border_style = zcl_excel_style_border=>c_border_thin.

      lo_style_blue = lo_excel->add_new_style( ).
      lo_style_blue->fill->fgcolor-rgb  = vl_cellcolor."<- Aquí se asigna el color de la celda de cabecera
      lo_style_blue->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
      lo_style_blue->borders->allborders    = lo_border_dark.
      lo_style_blue->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
      lv_style_blue_guid          = lo_style_blue->get_guid( ).

      lo_style_border = lo_excel->add_new_style( ).
      lo_style_border->borders->allborders = lo_border_dark.
      lv_style_border_guid = lo_style_border->get_guid( ).

      loop at i_libros assigning field-symbol(<fs_libros>).
            i_data_aux[] = i_data[].
            delete i_data_aux where indexlibro <> <fs_libros>-indexlibro.
            describe table i_data_aux lines vl_cantidad_reg.

            check i_data_aux[] is not initial.
            if sy-tabix eq 1.
               lo_worksheet = lo_excel->get_active_worksheet( ).
               lo_worksheet->set_title( ip_title = <fs_libros>-libro ).
            else.
               lo_worksheet = lo_excel->add_new_worksheet( ip_title = <fs_libros>-libro ).
            endif.

               go_table ?= cl_abap_typedescr=>describe_by_data( i_data ).
               go_struct ?= go_table->get_table_line_type( ).
               gt_comp = go_struct->get_components( ).

               vl_row = 1.

               do vl_cantidad_reg times.
               read table i_data_aux assigning field-symbol(<fs_data>) with key indexlibro = <fs_libros>-indexlibro."Se recorre la estructura libro por libro
               if sy-subrc eq 0.
                  vl_current_row = sy-tabix.

                  vl_letra = <fs_data>-columna.
                  vl_saltodelinea = <fs_data>-saltodelinea.

                  read table gt_comp assigning field-symbol(<fs_componente>) index 2 ." Se lee el valor de la columna
                  check sy-subrc eq 0.
                  concatenate '<fs_data>-' <FS_COMPONENTE>-NAME into vl_valor_fs.
                  assign (vl_valor_fs) to <fs_valor>.
                  move <fs_valor> to vl_value.

                  read table gt_comp assigning <fs_componente> index 4."Se determina si es una celda de cabecera
                  check sy-subrc eq 0.
                  concatenate '<fs_data>-' <FS_COMPONENTE>-NAME into vl_valor_fs.
                  assign (vl_valor_fs) to <fs_valor>.
                  if <fs_valor> eq abap_true.
                     read table gt_comp assigning <fs_componente> index 3."Se lee el valor que irá en la celda
                     check sy-subrc eq 0.
                     concatenate '<fs_data>-' <FS_COMPONENTE>-NAME into vl_valor_fs.
                     assign (vl_valor_fs) to <fs_valor>.
                     lo_worksheet->set_cell( ip_column = vl_letra ip_row = vl_row ip_value = <fs_valor> ip_style = lv_style_blue_guid ).
                  else.
                     read table gt_comp assigning <fs_componente> index 3."Se lee el valor que irá en la celda
                     check sy-subrc eq 0.
                     concatenate '<fs_data>-' <FS_COMPONENTE>-NAME into vl_valor_fs.
                     assign (vl_valor_fs) to <fs_valor>.
                     lo_worksheet->set_cell( ip_column = vl_letra ip_row = vl_row ip_value = <fs_valor> ).
                  endif.

                  lo_column = lo_worksheet->get_column( ip_column = vl_letra ).
                  lo_column->set_width( ip_width = 20 )."Se asigna el tamaño de la celda

                  delete i_data_aux index vl_current_row."Se borra la fila para no volverla a procesar
                  if vl_saltodelinea eq abap_true."Si hay salto de linea se continúa en la fila siguiente
                     vl_row = vl_row + 1.
                  endif.
               else.
                 continue.
               endif.
               enddo.
      endloop.

      if i_totales is not initial.
         loop at i_totales assigning field-symbol(<fs_totales>).
              condense <fs_totales>-valorcelda.
              lo_worksheet->set_cell( ip_column = <fs_totales>-columna ip_row = <fs_totales>-fila ip_value = <fs_totales>-valorcelda ).
         endloop.
      endif.

      if i_mergecells is not initial.
         loop at i_mergecells assigning field-symbol(<fs_mergecells>).
            if sy-subrc eq 0.
               condense <fs_mergecells>-valorcelda.
               lo_worksheet->set_cell( ip_row = <fs_mergecells>-indexmerge ip_column = <fs_mergecells>-inicio ip_value = <fs_mergecells>-valorcelda ip_style = lv_style_blue_guid ).
               lo_worksheet->set_merge( ip_row = <fs_mergecells>-indexmerge ip_column_start = <fs_mergecells>-inicio ip_column_end = <fs_mergecells>-fin ip_row_to = <fs_mergecells>-indexmerge ip_style = lv_style_blue_guid ).
            endif.
         endloop.
      endif.

      try.
          CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007."Se crea el documento
          vl_data = lo_excel_writer->write_file( lo_excel ).
          lt_rawdata = cl_bcs_convert=>xstring_to_solix( iv_xstring  = vl_data ).
      endtry.
  endtry.


"Se convierte a binario la data
CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
  EXPORTING
    buffer                = vl_data
 IMPORTING
   OUTPUT_LENGTH         = lv_bin_size
  tables
    binary_tab            = t_data_tab.

"Se descarga
 CALL FUNCTION 'GUI_DOWNLOAD'
 EXPORTING
 bin_filesize = lv_bin_size
 filename     = vl_filename
 filetype     = 'BIN'
 TABLES
   data_tab     = t_data_tab.


ENDFUNCTION.