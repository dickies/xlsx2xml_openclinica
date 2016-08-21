package xlsx2xml;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.GnuParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class biomedidas {
  static DecimalFormat df = new DecimalFormat("#####0");
  static SimpleDateFormat dateFormat = new SimpleDateFormat("YYYY-mm-dd");
  static SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
  
  public static void main(String[] args)
  {
    Boolean dbfunctions = Boolean.valueOf(false);
    
    PrintWriter out = null;
    String strOutputPath = "/tmp";
    String strFilePrefix = "importacion2";
    String StudyOID = "S_ATLANTIC";
    String SubjectKey = null;
    String StudySubjectID = "test";
    String StudyEventOID = "SE_TELEPAC";
    String FormOID = "F_BIOMEDIDAS_10";
    int crf_id = 15;
    int numeroSujetos = 0;
    String Item = null;
    String Igroup = null;
    String fecha_anterior = null;
    boolean date_registration = false;
    
    Date date_visit0 = null;
    Date date_visit1 = null;
    Date date_visit2 = null;
    Date date_visit3 = null;
    Date date_visit4 = null;
    String dateV0 = null;
    String dateV1 = null;
    String dateV2 = null;
    String dateV3 = null;
    String dateV4 = null;
    Date fecha_item = null;
    
    File inputFile = null;
    String outputFile = "/tmp/importacion.xml";
    String ddbb = "192.168.1.79";
    String port_ddbb = "5435";
    boolean first_import = false;
    
    Options options = new Options();
    options.addOption("h", false, "Mostrar ayuda");
    options.addOption("i", true, "Archivo input");
    options.addOption("o", true, "Archivo output");
    options.addOption("d", true, "Host database, value defaul openclinica.caebi.es");
    options.addOption("p", true, "Port database, value default 5433");
    options.addOption("f", false, "First import");
    options.addOption("c", true, "crf id");
    
    CommandLineParser cli = new GnuParser();
    try
    {
      CommandLine cmd = cli.parse(options, args);
      if (cmd.hasOption("h"))
      {
        printHelp(options);
        System.exit(0);
      }
      if (cmd.hasOption("d")) {
        ddbb = cmd.getOptionValue("d");
      }
      if (cmd.hasOption("p")) {
        port_ddbb = cmd.getOptionValue("p");
      }
      if (cmd.hasOption("f")) {
        first_import = true;
      }
      if (cmd.hasOption("c")) {
        crf_id = Integer.parseInt(cmd.getOptionValue("c"));
      }
      if (cmd.hasOption("i"))
      {
        try
        {
          inputFile = new File(cmd.getOptionValue("i"));
          if (!inputFile.exists())
          {
            System.err.println("El archivo no existe");
            System.exit(1);
          }
          if (!inputFile.isFile()) {
          System.err.println("El archivo no es valido");
          System.exit(1);
          }
        }
        catch (Exception ex)
        {
          System.err.println("El archivo no es valido");
          System.exit(1);
        }
      }
      else
      {
        System.err.println("El archivo es obligatorio");
        System.exit(1);
      }
      if (cmd.hasOption("o")) {
        try
        {
          outputFile = cmd.getOptionValue("o");
        }
        catch (Exception ex)
        {
          System.err.println("El archivo no es valido");
          System.exit(1);
        }
      }
    }
    catch (ParseException e1)
    {
      System.err.println("Invalid arguments");
      printHelp(options);
      System.exit(-1);
    }
    String driver = "org.postgresql.Driver";
    String connectString = "jdbc:postgresql://" + ddbb + ":" + port_ddbb + "/openclinica";
    String user = "clinica";
    String password = "clinica";
    try
    {
      InputStream inputStream = new FileInputStream(new File(inputFile.toString()));
      Workbook wb = WorkbookFactory.create(inputStream);
      Sheet sheet = wb.getSheet("data");
      
      FileWriter fostream = new FileWriter(outputFile);
      out = new PrintWriter(new BufferedWriter(fostream));
      try
      {
        Class.forName(driver);
        Connection con = DriverManager.getConnection(connectString, user, password);
        Statement stmt = con.createStatement();
        
        Date Date = new Date();
        SimpleDateFormat start = new SimpleDateFormat("YYYY-MM-dd HH:mm");
        System.out.println("Fecha: " + start.format(Date));
        System.out.println("Fichero de entrada " + inputFile.toString());
        System.out.println("Fichero de salida " + outputFile);
        out.println("<?xml version=\"1.0\" encoding=\"US-ASCII\"?><ODM xmlns=\"http://www.cdisc.org/ns/odm/v1.3\" xmlns:OpenClinica=\"http://www.openclinica.org/ns/odm_ext_v130/v3.1\" xmlns:OpenClinicaRules=\"http://www.openclinica.org/ns/rules/v3.1\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" FileOID=\"20151123.Biomedidas_filledD20151123161349+0100\" Description=\"20151123.Biomedidas_filled\" CreationDateTime=\"2015-11-23T16:13:49+01:00\" FileType=\"Snapshot\" ODMVersion=\"1.3\" xsi:schemaLocation=\"http://www.cdisc.org/ns/odm/v1.3 OpenClinica-ODM1-3-0-OC2-0.xsd\">");
        out.println("\t<ClinicalData StudyOID=\"" + StudyOID + "\" MetaDataVersionOID=\"v1.0.0\">");
        
        String parent_id = "";
        String biomedida = "";
        String original_biomedida = "";
        int controlador = 0;
        int repeatkey = 1;
        
        boolean firstRow = true;
        for (Row row : sheet) {
          if (firstRow)
          {
            firstRow = false;
          }
          else if (row.getRowNum() > controlador)
          {
            numeroSujetos++;
            int Subject_id = 0;
            int fila = 0;
            if ((!formatCell(row.getCell(0)).equals(parent_id)) || (!formatCell(row.getCell(5)).equals(original_biomedida)))
            {
              original_biomedida = formatCell(row.getCell(5));
              if (formatCell(row.getCell(5)).equals("TEMPERATURE"))
              {
                Igroup = "TEMPERATURA_3";
                Item = "TRA_3";
              }
              else if (formatCell(row.getCell(5)).equals("HEART RATE"))
              {
                Igroup = "SATURACION_3";
                Item = "SAT_3";
              }
              else if (formatCell(row.getCell(5)).equals("WEIGHT"))
              {
                Igroup = "PESO_3";
                Item = "PESO_3";
              }
              else if (formatCell(row.getCell(5)).equals("DIASTOLIC"))
              {
                Igroup = "TENSION_ARTERIAL_3";
                Item = "TA_3";
              }
              else if (formatCell(row.getCell(5)).equals("BASAL GLUCOSE LEVEL"))
              {
                Igroup = "GLUCOSA_BASAL_3";
                Item = "GLC_3";
              }
              else if (formatCell(row.getCell(5)).equals("GLUCOSE LEVEL BEFORE EATING"))
              {
                Igroup = "GLUCOSA_PREPRANDIAL_3";
                Item = "GLC_PRE_3";
              }
              else if (formatCell(row.getCell(5)).equals("GLUCOSE LEVEL AFTER EATING"))
              {
                Igroup = "GLUCOSA_POSTPRANDIAL_3";
                Item = "GLC_POST_3";
              }
              if (!formatCell(row.getCell(0)).equals(parent_id)) {
                date_registration = false;
              }
              parent_id = formatCell(row.getCell(0));
              biomedida = Igroup;
              repeatkey = 1;
              if (!first_import)
              {
                ResultSet rs2 = stmt.executeQuery("select study_subject_id from study_subject where oc_oid like 'SS_" + formatCell(row.getCell(0)) + "'");
                while (rs2.next()) {
                  Subject_id = Integer.parseInt(rs2.getString("study_subject_id"));
                }
                ResultSet rs = stmt.executeQuery("select item_data.ordinal from item_data, event_crf, crf, study_event,item_group, item_group_metadata where item_data.event_crf_id = event_crf.event_crf_id and event_crf.study_event_id = study_event.study_event_id and study_event.study_subject_id = " + Subject_id + " and item_data.item_id = item_group_metadata.item_id and item_group_metadata.item_group_id = item_group.item_group_id and item_group.crf_id = crf.crf_id and crf.crf_id = " + crf_id + " and item_group.oc_oid like 'IG_BIOME_" + Igroup + "' group by item_data.ordinal order by item_data.ordinal desc limit 1");
                while (rs.next()) {
                  if (Integer.parseInt(rs.getString("ordinal")) != 1) {
                    repeatkey = Integer.parseInt(rs.getString("ordinal")) + 1;
                  } else {
                    repeatkey = Integer.parseInt(rs.getString("ordinal"));
                  }
                }
              }
            }
            if (!date_registration)
            {
              ResultSet rs_v0 = stmt.executeQuery("select study_event.date_start from study_event,study_subject,study_event_definition where study_event_definition.name =  'Visita 0' and study_event.study_event_definition_id = study_event_definition.study_event_definition_id  and study_event.study_subject_id = study_subject.study_subject_id and study_event.study_subject_id =" + Subject_id);
              if (rs_v0.next())
              {
                date_visit0 = rs_v0.getDate("date_start");
                SimpleDateFormat datevisita = new SimpleDateFormat("YYYY-MM-dd");
                dateV0 = datevisita.format(date_visit0);
              }
              else
              {
                dateV0 = "";
              }
              ResultSet rs_v1 = stmt.executeQuery("select study_event.date_start from study_event,study_subject,study_event_definition where study_event_definition.name =  'Visita 1' and study_event.study_event_definition_id = study_event_definition.study_event_definition_id  and study_event.study_subject_id = study_subject.study_subject_id and study_event.study_subject_id =" + Subject_id);
              if (rs_v1.next())
              {
                date_visit1 = rs_v1.getDate("date_start");
                SimpleDateFormat datevisita = new SimpleDateFormat("YYYY-MM-dd");
                dateV1 = datevisita.format(date_visit1);
              }
              else
              {
                dateV1 = "";
              }
              ResultSet rs_v2 = stmt.executeQuery("select study_event.date_start from study_event,study_subject,study_event_definition where study_event_definition.name =  'Visita 2' and study_event.study_event_definition_id = study_event_definition.study_event_definition_id  and study_event.study_subject_id = study_subject.study_subject_id and study_event.study_subject_id =" + Subject_id);
              if (rs_v2.next())
              {
                date_visit2 = rs_v2.getDate("date_start");
                SimpleDateFormat datevisita = new SimpleDateFormat("YYYY-MM-dd");
                dateV2 = datevisita.format(date_visit2);
              }
              else
              {
                dateV2 = "";
              }
              ResultSet rs_v3 = stmt.executeQuery("select study_event.date_start from study_event,study_subject,study_event_definition where study_event_definition.name =  'Visita 3' and study_event.study_event_definition_id = study_event_definition.study_event_definition_id  and study_event.study_subject_id = study_subject.study_subject_id and study_event.study_subject_id =" + Subject_id);
              if (rs_v3.next())
              {
                date_visit3 = rs_v3.getDate("date_start");
                SimpleDateFormat datevisita = new SimpleDateFormat("YYYY-MM-dd");
                dateV3 = datevisita.format(date_visit3);
              }
              else
              {
                dateV3 = "";
              }
              ResultSet rs_v4 = stmt.executeQuery("select study_event.date_start from study_event,study_subject,study_event_definition where study_event_definition.name =  'Visita 4' and study_event.study_event_definition_id = study_event_definition.study_event_definition_id  and study_event.study_subject_id = study_subject.study_subject_id and study_event.study_subject_id =" + Subject_id);
              if (rs_v4.next())
              {
                date_visit4 = rs_v4.getDate("date_start");
                SimpleDateFormat datevisita = new SimpleDateFormat("YYYY-MM-dd");
                dateV4 = datevisita.format(date_visit4);
              }
              else
              {
                dateV4 = "";
              }
            }
            SimpleDateFormat datetemp = new SimpleDateFormat("YYYY-MM-dd");
            Date cellValue = row.getCell(2).getDateCellValue();
            String fecha = datetemp.format(cellValue);
            
            SimpleDateFormat comparacion = new SimpleDateFormat("yyyy-MM-dd");
            
            SimpleDateFormat datetemp2 = new SimpleDateFormat("HH:mm:ss");
            Date cellValue2 = row.getCell(3).getDateCellValue();
            String time = datetemp2.format(cellValue2);
            
            SubjectKey = formatCell(row.getCell(0));
            if (!date_registration)
            {
              out.println("\t\t<SubjectData SubjectKey=\"SS_" + SubjectKey + "\" OpenClinica:StudySubjectID=\"" + StudySubjectID + "\">");
              out.println("\t\t\t<StudyEventData StudyEventOID=\"" + StudyEventOID + "\">");
              out.println("\t\t\t\t<FormData FormOID=\"" + FormOID + "\" OpenClinica:Status=\"initial data entry\">");
              out.println("\t\t\t\t\t\t<ItemGroupData ItemGroupOID=\"IG_BIOME_UNGROUPED\" ItemGroupRepeatKey=\"1\" TransactionType=\"Insert\" >");
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_FECHA_V0_3\" Value=\"" + dateV0 + "\"/>");
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_FECHA_V1_3\" Value=\"" + dateV1 + "\"/>");
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_FECHA_V2_3\" Value=\"" + dateV2 + "\"/>");
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_FECHA_V3_3\" Value=\"" + dateV3 + "\"/>");
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_FECHA_V4_3\" Value=\"" + dateV4 + "\"/>");
              out.println("\t\t\t\t\t\t</ItemGroupData>");
              out.println("\t\t\t\t</FormData>");
              out.println("\t\t\t</StudyEventData>");
              out.println("\t\t</SubjectData>");
              date_registration = true;
            }
            out.println("\t\t<SubjectData SubjectKey=\"SS_" + SubjectKey + "\" OpenClinica:StudySubjectID=\"" + StudySubjectID + "\">");
            out.println("\t\t\t<StudyEventData StudyEventOID=\"" + StudyEventOID + "\">");
            out.println("\t\t\t\t<FormData FormOID=\"" + FormOID + "\" OpenClinica:Status=\"initial data entry\">");
            
            out.println("\t\t\t\t\t\t<ItemGroupData ItemGroupOID=\"IG_BIOME_" + Igroup + "\" " + "ItemGroupRepeatKey=\"" + repeatkey + "\" TransactionType=\"Insert\"" + ">");
            if ((!formatCell(row.getCell(5)).equals("GLUCOSE LEVEL BEFORE EATING")) && (!formatCell(row.getCell(5)).equals("GLUCOSE LEVEL AFTER EATING"))) {
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_TIPO_" + Item + "\" " + "Value= \"" + row.getCell(1) + "\"/>");
            }
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_FECHA_" + Item + "\" " + "Value= \"" + fecha + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_HORA_" + Item + "\" " + "Value= \"" + time + "\"/>");
            if (formatCell(row.getCell(5)).equals("DIASTOLIC"))
            {
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_TA_DIAS_3\" Value= \"" + row.getCell(6) + "\">");
              out.println("\t\t\t\t\t\t\t<MeasurementUnitRef MeasurementUnitOID= \"" + row.getCell(7) + "\"/>");
              out.println("\t\t\t\t\t\t\t</ItemData>");
              fecha_anterior = row.getCell(2).toString();
              row = sheet.getRow(row.getRowNum() + 1);
              if ((!fecha_anterior.equals(row.getCell(2).toString())) || (!formatCell(row.getCell(5)).equals("HEART RATE")))
              {
                System.out.println("Fecha " + fecha + " DIASLOTIC del sujeto " + parent_id + " no corresponde con fecha HEART RATE o no es HEART RATE el siguiente valor");
                System.exit(1);
              }
              else
              {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_PULSOTA_3\" Value= \"" + row.getCell(6) + "\">");
                out.println("\t\t\t\t\t\t\t<MeasurementUnitRef MeasurementUnitOID= \"" + row.getCell(7) + "\"/>");
                out.println("\t\t\t\t\t\t\t</ItemData>");
              }
              row = sheet.getRow(row.getRowNum() + 1);
              if ((!fecha_anterior.equals(row.getCell(2).toString())) || (!formatCell(row.getCell(5)).equals("SYSTOLIC")))
              {
                System.out.println("Fecha " + fecha + " HEART RATE del sujeto " + parent_id + " no corresponde con fecha SYSTOLIC o no es SYSTOLIC el siguiente valor");
                System.exit(1);
              }
              else
              {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_TA_SIS_3\" Value= \"" + row.getCell(6) + "\">");
              }
              controlador = row.getRowNum();
            }
            else if (formatCell(row.getCell(5)).equals("HEART RATE"))
            {
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_SAT_3\" Value= \"" + row.getCell(6) + "\">");
              out.println("\t\t\t\t\t\t\t<MeasurementUnitRef MeasurementUnitOID= \"" + row.getCell(7) + "\"/>");
              out.println("\t\t\t\t\t\t\t</ItemData>");
              fecha_anterior = row.getCell(2).toString();
              row = sheet.getRow(row.getRowNum() + 1);
              if ((!fecha_anterior.equals(row.getCell(2).toString())) || (!formatCell(row.getCell(5)).equals("SPO2")))
              {
                System.out.println("Fecha " + fecha + " HEART RATE del sujeto " + parent_id + " no corresponde con fecha SPO2 o no es SPO2 el siguiente valor");
                System.out.println("Fecha " + fecha_anterior + " distinta " + row.getCell(2).toString() + " puede que " + formatCell(row.getCell(5)) + " no sea SPO2");
                System.exit(1);
              }
              else
              {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_PULSOOX_3\" Value= \"" + row.getCell(6) + "\">");
                controlador = row.getRowNum();
              }
            }
            else if (formatCell(row.getCell(5)).equals("BASAL GLUCOSE LEVEL"))
            {
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_GLC_3\" Value= \"" + row.getCell(6) + "\">");
            }
            else if (formatCell(row.getCell(5)).equals("GLUCOSE LEVEL BEFORE EATING"))
            {
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_GLC_PRE_3\" Value= \"" + row.getCell(6) + "\">");
            }
            else if (formatCell(row.getCell(5)).equals("GLUCOSE LEVEL AFTER EATING"))
            {
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_GLC_POST_3\" Value= \"" + row.getCell(6) + "\">");
            }
            else
            {
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_BIOME_VALOR_" + Item + "\" " + "Value= \"" + row.getCell(6) + "\">");
            }
            out.println("\t\t\t\t\t\t\t<MeasurementUnitRef MeasurementUnitOID= \"" + row.getCell(7) + "\"/>");
            out.println("\t\t\t\t\t\t\t</ItemData>");
            
            out.println("\t\t\t\t\t\t</ItemGroupData>");
            out.println("\t\t\t\t</FormData>");
            out.println("\t\t\t</StudyEventData>");
            out.println("\t\t</SubjectData>");
            
            repeatkey++;
          }
        }
        stmt.close();
        con.close();
      }
      catch (Exception e)
      {
        System.out.println(e.getMessage());
      }
      out.println("\t</ClinicalData>");
      out.println("</ODM>");
      
      out.flush();
      out.close();
    }
    catch (Exception e)
    {
      e.printStackTrace();
    }
    System.out.println("Fin de parseo, han sido importados " + numeroSujetos + " sujetos");
  }
  
  private static String formatCell(Cell cell)
  {
    if (cell == null) {
      return "";
    }
    switch (cell.getCellType())
    {
    case 3: 
      return "";
    case 4: 
      return Boolean.toString(cell.getBooleanCellValue());
    case 5: 
      return "*error*";
    case 0: 
      return df.format(cell.getNumericCellValue());
    case 1: 
      return cell.getStringCellValue();
    }
    return "<unknown value>";
  }
  
  private static void printHelp(Options options)
  {
    HelpFormatter formatter = new HelpFormatter();
    formatter.printHelp("biomedidas.jar", options);
  }
}
