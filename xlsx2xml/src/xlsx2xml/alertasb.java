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

public class alertasb
{
  static DecimalFormat df = new DecimalFormat("#####0");
  static SimpleDateFormat dateFormat = new SimpleDateFormat("YYYY-mm-dd");
  static SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
  
  public static void main(String[] args)
  {
    PrintWriter out = null;
    String StudyOID = "S_ATLANTIC";
    String SubjectKey = null;
    String StudySubjectID = "test";
    String StudyEventOID = "SE_TELEPAC";
    String FormOID = "F_ALERTAS_NEW_10";
    int crf_id = 17;
    int numeroSujetos = 0;
    
    File inputFile = null;
    String outputFile = "/tmp/importacion_cuestionario.xml";
    String ddbb = "192.168.1.79";
    String port_ddbb = "5435";
    boolean first_import = false;
    
    Options options = new Options();
    options.addOption("h", false, "Mostrar ayuda");
    options.addOption("i", true, "Archivo input");
    options.addOption("o", true, "Archivo output");
    options.addOption("d", true, "Host database");
    options.addOption("p", true, "Port database");
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
        
        int repeatkey = 1;
        String parent_id = "";
        
        boolean firstRow = true;
        for (Row row : sheet) {
          if (firstRow)
          {
            firstRow = false;
          }
          else
          {
            numeroSujetos++;
            int Subject_id = 0;
            if (!formatCell(row.getCell(0)).equals(parent_id))
            {
              parent_id = formatCell(row.getCell(0));
              repeatkey = 1;
              if (!first_import)
              {
                ResultSet rs2 = stmt.executeQuery("select study_subject_id from study_subject where oc_oid like 'SS_" + formatCell(row.getCell(0)) + "'");
                while (rs2.next()) {
                  Subject_id = Integer.parseInt(rs2.getString("study_subject_id"));
                }
                ResultSet rs = stmt.executeQuery("select item_data.ordinal from item_data, event_crf, crf, study_event,item_group, item_group_metadata where item_data.event_crf_id = event_crf.event_crf_id and event_crf.study_event_id = study_event.study_event_id and study_event.study_subject_id = " + Subject_id + " and item_data.item_id = item_group_metadata.item_id and item_group_metadata.item_group_id = item_group.item_group_id and item_group.crf_id = crf.crf_id and crf.crf_id = " + crf_id + " and item_group.oc_oid like 'IG_ALERT_ALERTAS_N' group by item_data.ordinal order by item_data.ordinal desc limit 1");
                while (rs.next()) {
                  if (Integer.parseInt(rs.getString("ordinal")) != 1) {
                    repeatkey = Integer.parseInt(rs.getString("ordinal")) + 1;
                  } else {
                    repeatkey = Integer.parseInt(rs.getString("ordinal"));
                  }
                }
              }
            }
            SimpleDateFormat datetemp = new SimpleDateFormat("YYYY-MM-dd");
            Date cellValue = row.getCell(10).getDateCellValue();
            String fecha = datetemp.format(cellValue);
            
            SimpleDateFormat datetemp2 = new SimpleDateFormat("HH:mm:ss");
            Date cellValue2 = row.getCell(11).getDateCellValue();
            String time = datetemp2.format(cellValue2);
            
            Date cellValue3 = row.getCell(1).getDateCellValue();
            String nacimiento = datetemp.format(cellValue3);
            
            Date cellValue4 = row.getCell(4).getDateCellValue();
            String enrolment = datetemp.format(cellValue4);
            
            SubjectKey = formatCell(row.getCell(0));
            
            out.println("\t\t<SubjectData SubjectKey=\"SS_" + SubjectKey + "\" OpenClinica:StudySubjectID=\"" + StudySubjectID + "\">");
            out.println("\t\t\t<StudyEventData StudyEventOID=\"" + StudyEventOID + "\">");
            out.println("\t\t\t\t<FormData FormOID=\"" + FormOID + "\" OpenClinica:Status=\"initial data entry\">");
            out.println("\t\t\t\t\t\t<ItemGroupData ItemGroupOID=\"IG_ALERT_ALERTAS_N\" ItemGroupRepeatKey=\"" + repeatkey + "\" TransactionType=\"Insert\"" + ">");
            
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_BIRTH_DATE_N\" Value= \"" + nacimiento + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_PROGRAM_NAME_N\" Value= \"" + row.getCell(2) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_SEX_N\" Value= \"" + row.getCell(3) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_DATE_OF_ENROLMENT_N\" Value= \"" + enrolment + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_ORGANIZATION_N\" Value= \"" + row.getCell(5) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_CARE_PROVIDER_N\" Value= \"" + row.getCell(6) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_CALL_HANDLER_ID_N\" Value= \"" + row.getCell(7) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_CALL_HANDLER_SURNAME_N\" Value= \"" + row.getCell(8) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_CALL_HANDLER_FIRST_NAME_N\" Value= \"" + row.getCell(9) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_ALERT_DATE_N\" Value= \"" + fecha + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_ALERT_TIME_N\" Value= \"" + time + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_ALERT_ID_N\" Value= \"" + row.getCell(12) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_SEVERITY_N\" Value= \"" + row.getCell(13) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_THRESHOLD_N\" Value= \"" + row.getCell(14) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_ALERT_TYPE_N\" Value= \"" + row.getCell(15) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_ALERT_MSG_N\" Value= \"" + row.getCell(16) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_ALERT_STATUS_N\" Value= \"" + row.getCell(17) + "\"/>");
            out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_ALERT_CONTACTEME_N\" Value= \"" + row.getCell(18) + "\"/>");
            
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
      catch (Exception localException1) {}
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
  
  private static String formatElement(String prefix, String tag, String value)
  {
    StringBuilder sb = new StringBuilder(prefix);
    sb.append("<");
    sb.append(tag);
    if ((value != null) && (value.length() > 0))
    {
      sb.append(">");
      sb.append(value);
      sb.append("</");
      sb.append(tag);
      sb.append(">");
    }
    else
    {
      sb.append("/>");
    }
    return sb.toString();
  }
  
  private static void printHelp(Options options)
  {
    HelpFormatter formatter = new HelpFormatter();
    formatter.printHelp("parsebiomedidas.jar", options);
  }
}
