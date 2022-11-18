using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Drawing;
using ExcelUtil;

namespace AdressesUtility
{
    /// <summary>Utility functions for printing of addresses</summary>
    public class PrintJob
    {
        /// <summary>Input data for printing</summary>
        private PrintInput m_print_input;

        /// <summary>Input table with addresses</summary>
        private Table m_table_addresses = null;

        /// <summary>Class for the mapping of caption to field index</summary>
        private MapCaptionField m_map_caption_field;

        /// <summary>Print document used in the case that the user selects a printer</summary>
        private System.Drawing.Printing.PrintDocument m_print_document_default;

        //// <summary>A4 form width</summary>
        //private static float a4_width = 210.0f;

        //// <summary>A4 form height</summary>
        //private static float a4_height = 297.0f;

        //// <summary>Text font</summary>
        private Font m_text_font;

        public PrintJob(PrintInput i_print_input, Table i_table_addresses)
        {
            this.m_print_input = i_print_input;
            this.m_table_addresses = i_table_addresses;

            this.m_print_document_default = new System.Drawing.Printing.PrintDocument();
            this.m_print_document_default.DocumentName = "DefaultPrinterC#";
            this.m_print_document_default.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.m_print_document_default_PrintPage);
            this.m_print_document_default.EndPrint += new System.Drawing.Printing.PrintEventHandler(this.m_print_document_default_EndPrint);

            m_text_font = new Font(m_print_input.m_font, m_print_input.m_font_size);

            this.m_map_caption_field = new MapCaptionField(i_print_input, i_table_addresses);
        }

        /// <summary>Print with the default printer</summary>
        public bool DefaultPrinterPrint(out string o_error)
        {
            o_error = "";

            if (!m_map_caption_field.Map(out o_error))
                return false;

            m_print_document_default.Print();

            return true;
        }

        /// <summary>Sets sizes for fields and deltas between fields</summary>
        public bool SetSizesForFieldsAndDeltas(Graphics i_graphics, out string o_error)
        {
            o_error = "";

            _InitializeFieldSizes();

            // Note index start is one (0), i.e. also the header data is analyzed
            for (int i_row = 0; i_row < m_table_addresses.NumberRows; i_row++)
            {
                Row current_row = m_table_addresses.GetRow(i_row, out o_error);
                if (o_error != "") return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_10, ref m_print_input.m_size_mm_10, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_11, ref m_print_input.m_size_mm_11, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_12, ref m_print_input.m_size_mm_12, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_13, ref m_print_input.m_size_mm_13, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_20, ref m_print_input.m_size_mm_20, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_21, ref m_print_input.m_size_mm_21, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_22, ref m_print_input.m_size_mm_22, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_23, ref m_print_input.m_size_mm_23, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_30, ref m_print_input.m_size_mm_30, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_31, ref m_print_input.m_size_mm_31, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_32, ref m_print_input.m_size_mm_32, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_33, ref m_print_input.m_size_mm_33, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_40, ref m_print_input.m_size_mm_40, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_41, ref m_print_input.m_size_mm_41, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_42, ref m_print_input.m_size_mm_42, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_43, ref m_print_input.m_size_mm_43, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_50, ref m_print_input.m_size_mm_50, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_51, ref m_print_input.m_size_mm_51, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_52, ref m_print_input.m_size_mm_52, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_53, ref m_print_input.m_size_mm_53, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_60, ref m_print_input.m_size_mm_60, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_61, ref m_print_input.m_size_mm_61, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_62, ref m_print_input.m_size_mm_62, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_63, ref m_print_input.m_size_mm_63, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_70, ref m_print_input.m_size_mm_70, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_71, ref m_print_input.m_size_mm_71, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_72, ref m_print_input.m_size_mm_72, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_73, ref m_print_input.m_size_mm_73, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_80, ref m_print_input.m_size_mm_80, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_81, ref m_print_input.m_size_mm_81, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_82, ref m_print_input.m_size_mm_82, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_83, ref m_print_input.m_size_mm_83, ref m_print_input.m_text_height, out o_error)) return false;

                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_90, ref m_print_input.m_size_mm_90, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_91, ref m_print_input.m_size_mm_91, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_92, ref m_print_input.m_size_mm_92, ref m_print_input.m_text_height, out o_error)) return false;
                if (!_MaxWidthsForFieldsOneRow(i_graphics, current_row, m_print_input.m_caption_93, ref m_print_input.m_size_mm_93, ref m_print_input.m_text_height, out o_error)) return false;
            }

            return true;
        }

        /// <summary>Determine the text height</summary>
        private bool _MaxHeightsForFieldsOneRow(Graphics i_graphics, Row i_row, string i_caption, ref float io_height, out string o_error)
        {
            o_error = "";


            return true;
        }

        /// <summary>Maximum field sizes for one row</summary>
        private bool _MaxWidthsForFieldsOneRow(Graphics i_graphics, Row i_row, string i_caption, ref float io_width, ref float io_height, out string o_error)
        {
            o_error = "";

            if ("" == i_caption)
                return true; // Field that not shall be printed

            int field_index = m_map_caption_field.FieldIndex(i_caption);

            Field current_field = i_row.GetField(field_index, out o_error);
            if (o_error != "") 
                return false;

            string field_string = current_field.FieldValue;

            float field_width = _StringWidth(field_string, m_text_font, i_graphics);

            if (field_width > io_width)
            {
                io_width = field_width;
            }

            float field_height = _StringHeight(field_string, m_text_font, i_graphics);

            if (field_height > io_height)
            {
                io_height = field_height;
            }

            return true;
        }

        /// <summary>Initializes the field sizes to -1.0</summary>
        private void _InitializeFieldSizes()
        {
            this.m_print_input.m_size_mm_10 = -1.0f;
            this.m_print_input.m_size_mm_11 = -1.0f;
            this.m_print_input.m_size_mm_12 = -1.0f;
            this.m_print_input.m_size_mm_13 = -1.0f;

            this.m_print_input.m_size_mm_20 = -1.0f;
            this.m_print_input.m_size_mm_21 = -1.0f;
            this.m_print_input.m_size_mm_22 = -1.0f;
            this.m_print_input.m_size_mm_23 = -1.0f;

            this.m_print_input.m_size_mm_30 = -1.0f;
            this.m_print_input.m_size_mm_31 = -1.0f;
            this.m_print_input.m_size_mm_32 = -1.0f;
            this.m_print_input.m_size_mm_33 = -1.0f;

            this.m_print_input.m_size_mm_40 = -1.0f;
            this.m_print_input.m_size_mm_41 = -1.0f;
            this.m_print_input.m_size_mm_42 = -1.0f;
            this.m_print_input.m_size_mm_43 = -1.0f;

            this.m_print_input.m_size_mm_50 = -1.0f;
            this.m_print_input.m_size_mm_51 = -1.0f;
            this.m_print_input.m_size_mm_52 = -1.0f;
            this.m_print_input.m_size_mm_53 = -1.0f;

            this.m_print_input.m_size_mm_60 = -1.0f;
            this.m_print_input.m_size_mm_61 = -1.0f;
            this.m_print_input.m_size_mm_62 = -1.0f;
            this.m_print_input.m_size_mm_63 = -1.0f;

            this.m_print_input.m_size_mm_70 = -1.0f;
            this.m_print_input.m_size_mm_71 = -1.0f;
            this.m_print_input.m_size_mm_72 = -1.0f;
            this.m_print_input.m_size_mm_73 = -1.0f;

            this.m_print_input.m_size_mm_80 = -1.0f;
            this.m_print_input.m_size_mm_81 = -1.0f;
            this.m_print_input.m_size_mm_82 = -1.0f;
            this.m_print_input.m_size_mm_83 = -1.0f;

            this.m_print_input.m_size_mm_90 = -1.0f;
            this.m_print_input.m_size_mm_91 = -1.0f;
            this.m_print_input.m_size_mm_92 = -1.0f;
            this.m_print_input.m_size_mm_93 = -1.0f;

        }


        /// <summary>The print is ended. Nothing is done at the moment</summary>
        private void m_print_document_default_EndPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
        }

        /// <summary>Print page event. Printing data with the default printer</summary>
        private void m_print_document_default_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs i_event)
        {
            _SetSizesForFieldsAndDeltas(i_event.Graphics);

            _PrintAll(i_event.Graphics);
        }

        /// <summary>Send all data to the default printer</summary>
        private void _PrintAll(Graphics io_graphics)
        {
            io_graphics.Clear(System.Drawing.Color.White);
            io_graphics.PageUnit = GraphicsUnit.Millimeter;

            string error_message = "";

            // Note index start is one (0). Header data is also written
            for (int i_row = 0; i_row < m_table_addresses.NumberRows; i_row++)
            {
                Row current_row = m_table_addresses.GetRow(i_row, out error_message);
                if (error_message != "") return;

                _PrintRow(i_row, current_row, m_text_font, io_graphics);
            }


        }

        /// <summary>Set sizes for fields and deltas</summary>
        private void _SetSizesForFieldsAndDeltas(Graphics io_graphics)
        {
            string error_message = "";

            io_graphics.PageUnit = GraphicsUnit.Millimeter;

            if (!SetSizesForFieldsAndDeltas(io_graphics, out error_message))
                return;

        }

        /// <summary>Print row</summary>
        private void _PrintRow(int i_row_index, Row i_current_row, Font i_text_font, Graphics io_graphics)
        {
            int n_lines = 4;
            float row_y = m_print_input.m_marginal_top + (float)i_row_index * (float)n_lines * (m_print_input.m_text_height + 2.0f * m_print_input.m_delta_v);

           PointF position_10 = new PointF(m_print_input.m_marginal_left, row_y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_10, position_10, io_graphics);

           PointF position_11 = new PointF(position_10.X + m_print_input.m_size_mm_10 + m_print_input.m_delta_h, position_10.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_11, position_11, io_graphics);

           PointF position_12 = new PointF(position_11.X + m_print_input.m_size_mm_11 + m_print_input.m_delta_h, position_10.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_12, position_12, io_graphics);

           PointF position_13 = new PointF(position_12.X + m_print_input.m_size_mm_12 + m_print_input.m_delta_h, position_10.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_13, position_13, io_graphics);

           PointF position_20 = new PointF(m_print_input.m_marginal_left, position_10.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_20, position_20, io_graphics);

           PointF position_21 = new PointF(position_20.X + m_print_input.m_size_mm_20 + m_print_input.m_delta_h, position_20.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_21, position_21, io_graphics);

           PointF position_22 = new PointF(position_21.X + m_print_input.m_size_mm_21 + m_print_input.m_delta_h, position_20.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_22, position_22, io_graphics);

           PointF position_23 = new PointF(position_22.X + m_print_input.m_size_mm_22 + m_print_input.m_delta_h, position_20.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_23, position_23, io_graphics);

           PointF position_30 = new PointF(m_print_input.m_marginal_left, position_20.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_30, position_30, io_graphics);

           PointF position_31 = new PointF(position_30.X + m_print_input.m_size_mm_30 + m_print_input.m_delta_h, position_30.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_31, position_31, io_graphics);

           PointF position_32 = new PointF(position_31.X + m_print_input.m_size_mm_31 + m_print_input.m_delta_h, position_30.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_32, position_32, io_graphics);

           PointF position_33 = new PointF(position_32.X + m_print_input.m_size_mm_32 + m_print_input.m_delta_h, position_30.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_33, position_33, io_graphics);

           PointF position_40 = new PointF(m_print_input.m_marginal_left, position_30.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_40, position_40, io_graphics);

           PointF position_41 = new PointF(position_40.X + m_print_input.m_size_mm_40 + m_print_input.m_delta_h, position_40.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_31, position_31, io_graphics);

           PointF position_42 = new PointF(position_41.X + m_print_input.m_size_mm_41 + m_print_input.m_delta_h, position_40.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_42, position_42, io_graphics);

           PointF position_43 = new PointF(position_42.X + m_print_input.m_size_mm_42 + m_print_input.m_delta_h, position_40.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_43, position_43, io_graphics);

           PointF position_50 = new PointF(m_print_input.m_marginal_left, position_40.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_50, position_50, io_graphics);

           PointF position_51 = new PointF(position_50.X + m_print_input.m_size_mm_50 + m_print_input.m_delta_h, position_50.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_51, position_51, io_graphics);

           PointF position_52 = new PointF(position_51.X + m_print_input.m_size_mm_51 + m_print_input.m_delta_h, position_50.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_52, position_52, io_graphics);

           PointF position_53 = new PointF(position_52.X + m_print_input.m_size_mm_52 + m_print_input.m_delta_h, position_50.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_53, position_53, io_graphics);

           PointF position_60 = new PointF(m_print_input.m_marginal_left, position_50.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_60, position_60, io_graphics);

           PointF position_61 = new PointF(position_60.X + m_print_input.m_size_mm_60 + m_print_input.m_delta_h, position_60.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_61, position_61, io_graphics);

           PointF position_62 = new PointF(position_61.X + m_print_input.m_size_mm_61 + m_print_input.m_delta_h, position_60.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_62, position_62, io_graphics);

           PointF position_63 = new PointF(position_62.X + m_print_input.m_size_mm_62 + m_print_input.m_delta_h, position_60.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_63, position_63, io_graphics);

           PointF position_70 = new PointF(m_print_input.m_marginal_left, position_60.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_70, position_70, io_graphics);

           PointF position_71 = new PointF(position_70.X + m_print_input.m_size_mm_70 + m_print_input.m_delta_h, position_70.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_71, position_71, io_graphics);

           PointF position_72 = new PointF(position_71.X + m_print_input.m_size_mm_71 + m_print_input.m_delta_h, position_70.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_72, position_72, io_graphics);

           PointF position_73 = new PointF(position_72.X + m_print_input.m_size_mm_72 + m_print_input.m_delta_h, position_70.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_73, position_73, io_graphics);

           PointF position_80 = new PointF(m_print_input.m_marginal_left, position_70.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_80, position_80, io_graphics);

           PointF position_81 = new PointF(position_80.X + m_print_input.m_size_mm_80 + m_print_input.m_delta_h, position_80.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_81, position_81, io_graphics);

           PointF position_82 = new PointF(position_81.X + m_print_input.m_size_mm_81 + m_print_input.m_delta_h, position_80.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_82, position_82, io_graphics);

           PointF position_83 = new PointF(position_82.X + m_print_input.m_size_mm_82 + m_print_input.m_delta_h, position_80.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_83, position_83, io_graphics);

           PointF position_90 = new PointF(m_print_input.m_marginal_left, position_80.Y + m_print_input.m_delta_v + m_print_input.m_text_height);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_90, position_90, io_graphics);

           PointF position_91 = new PointF(position_90.X + m_print_input.m_size_mm_90 + m_print_input.m_delta_h, position_90.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_91, position_91, io_graphics);

           PointF position_92 = new PointF(position_91.X + m_print_input.m_size_mm_91 + m_print_input.m_delta_h, position_90.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_92, position_92, io_graphics);

           PointF position_93 = new PointF(position_92.X + m_print_input.m_size_mm_92 + m_print_input.m_delta_h, position_90.Y);
           _PrintField(i_current_row, i_text_font, m_print_input.m_caption_93, position_93, io_graphics);

        }

        /// <summary>Print field</summary>
        private void _PrintField(Row i_current_row, Font i_text_font, string i_caption, PointF i_position, Graphics io_graphics)
        {

            int index_column = m_map_caption_field.FieldIndex(i_caption);

            if (index_column >= 0)
            {
                string error_message = "";

                Field current_field = i_current_row.GetField(index_column, out error_message);
                if (error_message != "") return;

                string current_value = current_field.FieldValue;

                io_graphics.DrawString(current_value, i_text_font, Brushes.Black, i_position);
            }
        }

        /// <summary>Send data to the default printer</summary>
        public void DefaultPrinterPrint()
        {
            m_print_document_default.Print();
        }

        /// <summary>Returns the width of a string with a given font in millimeter</summary>
        private float _StringWidth(string i_string, Font i_font, Graphics i_graphics)
        {
            float ret_length = -1.0f;

            SizeF string_size = new SizeF();
            string_size = i_graphics.MeasureString(i_string, i_font);

            ret_length = string_size.Width;

            return ret_length;
        }

        /// <summary>Returns the height of a string with a given font in millimeter</summary>
        private float _StringHeight(string i_string, Font i_font, Graphics i_graphics)
        {
            float ret_length = -1.0f;

            SizeF string_size = new SizeF();
            string_size = i_graphics.MeasureString(i_string, i_font);

            ret_length = string_size.Height;

            return ret_length;
        }
    }

    /// <summary>Input data for printing</summary>
    public class PrintInput
    {
        /// <summary>Font</summary>
        public string m_font = "Arial";

        /// <summary>Font size</summary>
        public float m_font_size = 6.0f;

        /// <summary>Text height in millimeter</summary>
        public float m_text_height = -1.0f;

        /// <summary>Marginal left in millimeter</summary>
        public float m_marginal_left = 10.0f;

        /// <summary>Marginal right in millimeter</summary>
        public float m_marginal_right = 10.0f;

        /// <summary>Marginal top in millimeter</summary>
        public float m_marginal_top = 10.0f;

        /// <summary>Marginal bottom in millimeter</summary>
        public float m_marginal_bottom = 10.0f;

        /// <summary>Delta horizontal in millimeter</summary>
        public float m_delta_h = 1.0f;

        /// <summary>Delta vertical in millimeter</summary>
        public float m_delta_v = 1.0f;

        /// <summary>Caption for data 10</summary>
        public string m_caption_10 = "";

        /// <summary>Size in millimeter for data 10</summary>
        public float m_size_mm_10 = 32.0f;

        /// <summary>Text prior to data 10</summary>
        public string m_text_10 = "";

        /// <summary>Caption for data 11</summary>
        public string m_caption_11 = "";

        /// <summary>Size in millimeter for data 11</summary>
        public float m_size_mm_11 = 32.0f;

        /// <summary>Text prior to data 11</summary>
        public string m_text_11 = "";

        /// <summary>Caption for data 12</summary>
        public string m_caption_12 = "";

        /// <summary>Size in millimeter for data 12</summary>
        public float m_size_mm_12 = 32.0f;

        /// <summary>Text prior to data 12</summary>
        public string m_text_12 = "";

        /// <summary>Caption for data 13</summary>
        public string m_caption_13 = "";

        /// <summary>Size in millimeter for data 13</summary>
        public float m_size_mm_13 = 32.0f;

        /// <summary>Text prior to data 13</summary>
        public string m_text_13 = "";

        /// <summary>Caption for data 20</summary>
        public string m_caption_20 = "";

        /// <summary>Size in millimeter for data 20</summary>
        public float m_size_mm_20 = 32.0f;

        /// <summary>Text prior to data 20</summary>
        public string m_text_20 = "";

        /// <summary>Caption for data 21</summary>
        public string m_caption_21 = "";

        /// <summary>Size in millimeter for data 21</summary>
        public float m_size_mm_21 = 32.0f;

        /// <summary>Text prior to data 21</summary>
        public string m_text_21 = "";

        /// <summary>Caption for data 22</summary>
        public string m_caption_22 = "";

        /// <summary>Size in millimeter for data 22</summary>
        public float m_size_mm_22 = 32.0f;

        /// <summary>Text prior to data 22</summary>
        public string m_text_22 = "";

        /// <summary>Caption for data 23</summary>
        public string m_caption_23 = "";

        /// <summary>Size in millimeter for data 23</summary>
        public float m_size_mm_23 = 32.0f;

        /// <summary>Text prior to data 23</summary>
        public string m_text_23 = "";

        /// <summary>Caption for data 30</summary>
        public string m_caption_30 = "";

        /// <summary>Size in millimeter for data 30</summary>
        public float m_size_mm_30 = 32.0f;

        /// <summary>Text prior to data 30</summary>
        public string m_text_30 = "";

        /// <summary>Caption for data 31</summary>
        public string m_caption_31 = "";

        /// <summary>Size in millimeter for data 31</summary>
        public float m_size_mm_31 = 32.0f;

        /// <summary>Text prior to data 31</summary>
        public string m_text_31 = "";

        /// <summary>Caption for data 32</summary>
        public string m_caption_32 = "";

        /// <summary>Size in millimeter for data 32</summary>
        public float m_size_mm_32 = 32.0f;

        /// <summary>Text prior to data 32</summary>
        public string m_text_32 = "";

        /// <summary>Caption for data 33</summary>
        public string m_caption_33 = "";

        /// <summary>Size in millimeter for data 33</summary>
        public float m_size_mm_33 = 32.0f;

        /// <summary>Text prior to data 33</summary>
        public string m_text_33 = "";

        /// <summary>Caption for data 40</summary>
        public string m_caption_40 = "";

        /// <summary>Size in millimeter for data 40</summary>
        public float m_size_mm_40 = 32.0f;

        /// <summary>Text prior to data 40</summary>
        public string m_text_40 = "";

        /// <summary>Caption for data 41</summary>
        public string m_caption_41 = "";

        /// <summary>Size in millimeter for data 41</summary>
        public float m_size_mm_41 = 32.0f;

        /// <summary>Text prior to data 41</summary>
        public string m_text_41 = "";

        /// <summary>Caption for data 42</summary>
        public string m_caption_42 = "";

        /// <summary>Size in millimeter for data 42</summary>
        public float m_size_mm_42 = 32.0f;

        /// <summary>Text prior to data 42</summary>
        public string m_text_42 = "";

        /// <summary>Caption for data 43</summary>
        public string m_caption_43 = "";

        /// <summary>Size in millimeter for data 43</summary>
        public float m_size_mm_43 = 32.0f;

        /// <summary>Text prior to data 43</summary>
        public string m_text_43 = "";

        /// <summary>Caption for data 50</summary>
        public string m_caption_50 = "";

        /// <summary>Size in millimeter for data 50</summary>
        public float m_size_mm_50 = 32.0f;

        /// <summary>Text prior to data 50</summary>
        public string m_text_50 = "";

        /// <summary>Caption for data 51</summary>
        public string m_caption_51 = "";

        /// <summary>Size in millimeter for data 51</summary>
        public float m_size_mm_51 = 32.0f;

        /// <summary>Text prior to data 51</summary>
        public string m_text_51 = "";

        /// <summary>Caption for data 52</summary>
        public string m_caption_52 = "";

        /// <summary>Size in millimeter for data 52</summary>
        public float m_size_mm_52 = 32.0f;

        /// <summary>Text prior to data 52</summary>
        public string m_text_52 = "";

        /// <summary>Caption for data 53</summary>
        public string m_caption_53 = "";

        /// <summary>Size in millimeter for data 53</summary>
        public float m_size_mm_53 = 32.0f;

        /// <summary>Text prior to data 53</summary>
        public string m_text_53 = "";

        /// <summary>Caption for data 60</summary>
        public string m_caption_60 = "";

        /// <summary>Size in millimeter for data 60</summary>
        public float m_size_mm_60 = 32.0f;

        /// <summary>Text prior to data 60</summary>
        public string m_text_60 = "";

        /// <summary>Caption for data 51</summary>
        public string m_caption_61 = "";

        /// <summary>Size in millimeter for data 61</summary>
        public float m_size_mm_61 = 32.0f;

        /// <summary>Text prior to data 61</summary>
        public string m_text_61 = "";

        /// <summary>Caption for data 62</summary>
        public string m_caption_62 = "";

        /// <summary>Size in millimeter for data 62</summary>
        public float m_size_mm_62 = 32.0f;

        /// <summary>Text prior to data 62</summary>
        public string m_text_62 = "";

        /// <summary>Caption for data 63</summary>
        public string m_caption_63 = "";

        /// <summary>Size in millimeter for data 63</summary>
        public float m_size_mm_63 = 32.0f;

        /// <summary>Text prior to data 63</summary>
        public string m_text_63 = "";

        /// <summary>Caption for data 70</summary>
        public string m_caption_70 = "";

        /// <summary>Size in millimeter for data 70</summary>
        public float m_size_mm_70 = 32.0f;

        /// <summary>Text prior to data 70</summary>
        public string m_text_70 = "";

        /// <summary>Caption for data 71</summary>
        public string m_caption_71 = "";

        /// <summary>Size in millimeter for data 71</summary>
        public float m_size_mm_71 = 32.0f;

        /// <summary>Text prior to data 71</summary>
        public string m_text_71 = "";

        /// <summary>Caption for data 72</summary>
        public string m_caption_72 = "";

        /// <summary>Size in millimeter for data 72</summary>
        public float m_size_mm_72 = 32.0f;

        /// <summary>Text prior to data 72</summary>
        public string m_text_72 = "";

        /// <summary>Caption for data 73</summary>
        public string m_caption_73 = "";

        /// <summary>Size in millimeter for data 73</summary>
        public float m_size_mm_73 = 32.0f;

        /// <summary>Text prior to data 73</summary>
        public string m_text_73 = "";

        /// <summary>Caption for data 80</summary>
        public string m_caption_80 = "";

        /// <summary>Size in millimeter for data 80</summary>
        public float m_size_mm_80 = 32.0f;

        /// <summary>Text prior to data 80</summary>
        public string m_text_80 = "";

        /// <summary>Caption for data 81</summary>
        public string m_caption_81 = "";

        /// <summary>Size in millimeter for data 81</summary>
        public float m_size_mm_81 = 32.0f;

        /// <summary>Text prior to data 81</summary>
        public string m_text_81 = "";

        /// <summary>Caption for data 82</summary>
        public string m_caption_82 = "";

        /// <summary>Size in millimeter for data 82</summary>
        public float m_size_mm_82 = 32.0f;

        /// <summary>Text prior to data 82</summary>
        public string m_text_82 = "";

        /// <summary>Caption for data 83</summary>
        public string m_caption_83 = "";

        /// <summary>Size in millimeter for data 83</summary>
        public float m_size_mm_83 = 32.0f;

        /// <summary>Text prior to data 83</summary>
        public string m_text_83 = "";

        /// <summary>Caption for data 90</summary>
        public string m_caption_90 = "";

        /// <summary>Size in millimeter for data 90</summary>
        public float m_size_mm_90 = 32.0f;

        /// <summary>Text prior to data 90</summary>
        public string m_text_90 = "";

        /// <summary>Caption for data 91</summary>
        public string m_caption_91 = "";

        /// <summary>Size in millimeter for data 91</summary>
        public float m_size_mm_91 = 32.0f;

        /// <summary>Text prior to data 91</summary>
        public string m_text_91 = "";

        /// <summary>Caption for data 92</summary>
        public string m_caption_92 = "";

        /// <summary>Size in millimeter for data 92</summary>
        public float m_size_mm_92 = 32.0f;

        /// <summary>Text prior to data 92</summary>
        public string m_text_92 = "";

        /// <summary>Caption for data 93</summary>
        public string m_caption_93 = "";

        /// <summary>Size in millimeter for data 93</summary>
        public float m_size_mm_93 = 32.0f;

        /// <summary>Text prior to data 93</summary>
        public string m_text_93 = "";

    }

    /// <summary>Map caption to field index</summary>
    public class MapCaptionField
    {
        /// <summary>Row header with captions</summary>
        RowHeader m_row_header;

        //private int m_fields =
        private int m_field_10 = -1;
        private int m_field_11 = -1;
        private int m_field_12 = -1;
        private int m_field_13 = -1;

        private int m_field_20 = -1;
        private int m_field_21 = -1;
        private int m_field_22 = -1;
        private int m_field_23 = -1;

        private int m_field_30 = -1;
        private int m_field_31 = -1;
        private int m_field_32 = -1;
        private int m_field_33 = -1;

        private int m_field_40 = -1;
        private int m_field_41 = -1;
        private int m_field_42 = -1;
        private int m_field_43 = -1;

        private int m_field_50 = -1;
        private int m_field_51 = -1;
        private int m_field_52 = -1;
        private int m_field_53 = -1;

        private int m_field_60 = -1;
        private int m_field_61 = -1;
        private int m_field_62 = -1;
        private int m_field_63 = -1;

        private int m_field_70 = -1;
        private int m_field_71 = -1;
        private int m_field_72 = -1;
        private int m_field_73 = -1;

        private int m_field_80 = -1;
        private int m_field_81 = -1;
        private int m_field_82 = -1;
        private int m_field_83 = -1;

        private int m_field_90 = -1;
        private int m_field_91 = -1;
        private int m_field_92 = -1;
        private int m_field_93 = -1;

        /// <summary>Input data for printing</summary>
        private PrintInput m_print_input;

		/// <summary>Input table with addresses</summary>
        private Table m_table_addresses = null;

        public MapCaptionField(PrintInput i_print_input, Table i_table_addresses)
        {
            this.m_print_input = i_print_input;
            this.m_table_addresses = i_table_addresses;
            this.m_row_header = m_table_addresses.GetRowHeader();
 
        }

        /// <summary>Map captions to field index</summary>
        public bool Map(out string o_error)
        {
            o_error = "";

            if (!_MapOneField(out this.m_field_10, this.m_print_input.m_caption_10, out o_error)) return false;
            if (!_MapOneField(out this.m_field_11, this.m_print_input.m_caption_11, out o_error)) return false;
            if (!_MapOneField(out this.m_field_12, this.m_print_input.m_caption_12, out o_error)) return false;
            if (!_MapOneField(out this.m_field_13, this.m_print_input.m_caption_13, out o_error)) return false;

            if (!_MapOneField(out this.m_field_20, this.m_print_input.m_caption_20, out o_error)) return false;
            if (!_MapOneField(out this.m_field_21, this.m_print_input.m_caption_21, out o_error)) return false;
            if (!_MapOneField(out this.m_field_22, this.m_print_input.m_caption_22, out o_error)) return false;
            if (!_MapOneField(out this.m_field_23, this.m_print_input.m_caption_23, out o_error)) return false;

            if (!_MapOneField(out this.m_field_30, this.m_print_input.m_caption_30, out o_error)) return false;
            if (!_MapOneField(out this.m_field_31, this.m_print_input.m_caption_31, out o_error)) return false;
            if (!_MapOneField(out this.m_field_32, this.m_print_input.m_caption_32, out o_error)) return false;
            if (!_MapOneField(out this.m_field_33, this.m_print_input.m_caption_33, out o_error)) return false;

            if (!_MapOneField(out this.m_field_40, this.m_print_input.m_caption_40, out o_error)) return false;
            if (!_MapOneField(out this.m_field_41, this.m_print_input.m_caption_41, out o_error)) return false;
            if (!_MapOneField(out this.m_field_42, this.m_print_input.m_caption_42, out o_error)) return false;
            if (!_MapOneField(out this.m_field_43, this.m_print_input.m_caption_43, out o_error)) return false;

            if (!_MapOneField(out this.m_field_50, this.m_print_input.m_caption_50, out o_error)) return false;
            if (!_MapOneField(out this.m_field_51, this.m_print_input.m_caption_51, out o_error)) return false;
            if (!_MapOneField(out this.m_field_52, this.m_print_input.m_caption_52, out o_error)) return false;
            if (!_MapOneField(out this.m_field_53, this.m_print_input.m_caption_53, out o_error)) return false;

            if (!_MapOneField(out this.m_field_60, this.m_print_input.m_caption_60, out o_error)) return false;
            if (!_MapOneField(out this.m_field_61, this.m_print_input.m_caption_61, out o_error)) return false;
            if (!_MapOneField(out this.m_field_62, this.m_print_input.m_caption_62, out o_error)) return false;
            if (!_MapOneField(out this.m_field_63, this.m_print_input.m_caption_63, out o_error)) return false;

            if (!_MapOneField(out this.m_field_70, this.m_print_input.m_caption_70, out o_error)) return false;
            if (!_MapOneField(out this.m_field_71, this.m_print_input.m_caption_71, out o_error)) return false;
            if (!_MapOneField(out this.m_field_72, this.m_print_input.m_caption_72, out o_error)) return false;
            if (!_MapOneField(out this.m_field_73, this.m_print_input.m_caption_73, out o_error)) return false;

            if (!_MapOneField(out this.m_field_80, this.m_print_input.m_caption_80, out o_error)) return false;
            if (!_MapOneField(out this.m_field_81, this.m_print_input.m_caption_81, out o_error)) return false;
            if (!_MapOneField(out this.m_field_82, this.m_print_input.m_caption_82, out o_error)) return false;
            if (!_MapOneField(out this.m_field_83, this.m_print_input.m_caption_83, out o_error)) return false;

            if (!_MapOneField(out this.m_field_90, this.m_print_input.m_caption_90, out o_error)) return false;
            if (!_MapOneField(out this.m_field_91, this.m_print_input.m_caption_91, out o_error)) return false;
            if (!_MapOneField(out this.m_field_92, this.m_print_input.m_caption_92, out o_error)) return false;
            if (!_MapOneField(out this.m_field_93, this.m_print_input.m_caption_93, out o_error)) return false;

            return true;
        }

        /// <summary>Map one field</summary>
        private bool _MapOneField(out int o_field_xy, string i_caption, out string o_error)
        {
            o_error = "";
            o_field_xy = -1;

            for (int i_column = 0; i_column < m_row_header.NumberColumns; i_column++)
            {
                FieldHeader field_header = m_row_header.GetFieldHeader(i_column, out o_error);
                if ("" != o_error) return false;

                if (field_header.Caption == i_caption)
                {
                    o_field_xy = i_column;
                    break;
                }

            }

            return true;
        }

        /// <summary>Returns the field index for a given caption</summary>
        public int FieldIndex(string i_caption)
        {
            if      (i_caption == m_print_input.m_caption_10) return m_field_10;
            else if (i_caption == m_print_input.m_caption_11) return m_field_11;
            else if (i_caption == m_print_input.m_caption_12) return m_field_12;
            else if (i_caption == m_print_input.m_caption_13) return m_field_13;

            else if (i_caption == m_print_input.m_caption_20) return m_field_20;
            else if (i_caption == m_print_input.m_caption_21) return m_field_21;
            else if (i_caption == m_print_input.m_caption_22) return m_field_22;
            else if (i_caption == m_print_input.m_caption_23) return m_field_23;

            else if (i_caption == m_print_input.m_caption_30) return m_field_30;
            else if (i_caption == m_print_input.m_caption_31) return m_field_31;
            else if (i_caption == m_print_input.m_caption_32) return m_field_32;
            else if (i_caption == m_print_input.m_caption_33) return m_field_33;

            else if (i_caption == m_print_input.m_caption_40) return m_field_40;
            else if (i_caption == m_print_input.m_caption_41) return m_field_41;
            else if (i_caption == m_print_input.m_caption_42) return m_field_42;
            else if (i_caption == m_print_input.m_caption_43) return m_field_43;

            else if (i_caption == m_print_input.m_caption_50) return m_field_50;
            else if (i_caption == m_print_input.m_caption_51) return m_field_51;
            else if (i_caption == m_print_input.m_caption_52) return m_field_52;
            else if (i_caption == m_print_input.m_caption_53) return m_field_53;

            else if (i_caption == m_print_input.m_caption_60) return m_field_60;
            else if (i_caption == m_print_input.m_caption_61) return m_field_61;
            else if (i_caption == m_print_input.m_caption_62) return m_field_62;
            else if (i_caption == m_print_input.m_caption_63) return m_field_63;

            else if (i_caption == m_print_input.m_caption_70) return m_field_70;
            else if (i_caption == m_print_input.m_caption_71) return m_field_71;
            else if (i_caption == m_print_input.m_caption_72) return m_field_72;
            else if (i_caption == m_print_input.m_caption_73) return m_field_73;

            else if (i_caption == m_print_input.m_caption_80) return m_field_80;
            else if (i_caption == m_print_input.m_caption_81) return m_field_81;
            else if (i_caption == m_print_input.m_caption_82) return m_field_82;
            else if (i_caption == m_print_input.m_caption_83) return m_field_83;

            else if (i_caption == m_print_input.m_caption_90) return m_field_90;
            else if (i_caption == m_print_input.m_caption_91) return m_field_91;
            else if (i_caption == m_print_input.m_caption_92) return m_field_92;
            else if (i_caption == m_print_input.m_caption_93) return m_field_93;


            else return -1;
        }

    }
}
