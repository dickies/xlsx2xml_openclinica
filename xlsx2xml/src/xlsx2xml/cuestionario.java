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

public class cuestionario
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
    String FormOID = "F_CUESTIONARIO_10";
    int crf_id = 16;
    int numeroSujetos = 0;
    String Item = null;
    String Igroup = null;
    String fecha_anterior = null;
    String tiempo_anterior = null;
    boolean primera_pasada = true;
    String score_global = "";
    String Igroup_anterior = null;
    
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
      label364:
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
        String original_organo = "";
        int repeatkey = 1;
        boolean primer = true;
        
        boolean firstRow = true;
        for (Row row : sheet) {
          if (firstRow)
          {
            firstRow = false;
          }
          else
          {
            if (1 == row.getRowNum())
            {
              SimpleDateFormat datetemp3 = new SimpleDateFormat("YYYY-MM-dd");
              Date cellValue3 = row.getCell(3).getDateCellValue();
              fecha_anterior = datetemp3.format(cellValue3);
              
              SimpleDateFormat datetemp4 = new SimpleDateFormat("HH:mm:ss");
              Date cellValue4 = row.getCell(4).getDateCellValue();
              tiempo_anterior = datetemp4.format(cellValue4);
            }
            numeroSujetos++;
            int Subject_id = 0;
            if ((!formatCell(row.getCell(0)).equals(parent_id)) || (!formatCell(row.getCell(2)).equals(original_organo)))
            {
              repeatkey = 0;
              if (primera_pasada)
              {
                repeatkey = 1;
                //primera_pasada = false;
              }
              original_organo = formatCell(row.getCell(2));
              if (formatCell(row.getCell(2)).equals("¿Cómo está mi corazón?"))
              {
                Igroup = "CORAZON";
                Item = "COR";
              }
              else if (formatCell(row.getCell(2)).equals("¿Cómo están mis pulmones?"))
              {
                Igroup = "PULMONES";
                Item = "PUL";
              }
              parent_id = formatCell(row.getCell(0));
              if (!first_import)
              {
                ResultSet rs2 = stmt.executeQuery("select study_subject_id from study_subject where oc_oid like 'SS_" + formatCell(row.getCell(0)) + "'");
                while (rs2.next()) {
                  Subject_id = Integer.parseInt(rs2.getString("study_subject_id"));
                }
                ResultSet rs = stmt.executeQuery("select item_data.ordinal from item_data, event_crf, crf, study_event,item_group, item_group_metadata where item_data.event_crf_id = event_crf.event_crf_id and event_crf.study_event_id = study_event.study_event_id and study_event.study_subject_id = " + Subject_id + " and item_data.item_id = item_group_metadata.item_id and item_group_metadata.item_group_id = item_group.item_group_id and item_group.crf_id = crf.crf_id and crf.crf_id = " + crf_id + " and item_group.oc_oid like 'IG_CUEST_" + Igroup + "' group by item_data.ordinal order by item_data.ordinal desc limit 1");
                while (rs.next()) {
                  if ((Integer.parseInt(rs.getString("ordinal")) != 1) && (!Igroup.equals("PULMONES")) && primera_pasada ) {
                    repeatkey = Integer.parseInt(rs.getString("ordinal")) + 1;
                    primera_pasada = false;
                  } else {
                    repeatkey = Integer.parseInt(rs.getString("ordinal"));
                  }
                }
              }
            }
            SimpleDateFormat datetemp = new SimpleDateFormat("YYYY-MM-dd");
            Date cellValue = row.getCell(3).getDateCellValue();
            String fecha = datetemp.format(cellValue);
            
            SimpleDateFormat datetemp2 = new SimpleDateFormat("HH:mm:ss");
            Date cellValue2 = row.getCell(4).getDateCellValue();
            String time = datetemp2.format(cellValue2);
            if ((!fecha.equals(fecha_anterior)) || (!time.equals(tiempo_anterior)))
            {
              if (Igroup != Igroup_anterior)
              {
                if (Igroup.equals("CORAZON")) {
                  out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_SCORE_PUL\" Value= \"" + score_global + "\"/>");
                } else if (Igroup.equals("PULMONES")) {
                  out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_SCORE_COR\" Value= \"" + score_global + "\"/>");
                }
              }
              else if (Igroup.equals("CORAZON")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_SCORE_COR\" Value= \"" + score_global + "\"/>");
              } else if (Igroup.equals("PULMONES")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_SCORE_PUL\" Value= \"" + score_global + "\"/>");
              }
              out.println("\t\t\t\t\t\t</ItemGroupData>");
              out.println("\t\t\t\t</FormData>");
              out.println("\t\t\t</StudyEventData>");
              out.println("\t\t</SubjectData>");
              primer = true;
              repeatkey++;
            }
            SubjectKey = formatCell(row.getCell(0));
            if (primer)
            {
              out.println("\t\t<SubjectData SubjectKey=\"SS_" + SubjectKey + "\" OpenClinica:StudySubjectID=\"" + StudySubjectID + "\">");
              out.println("\t\t\t<StudyEventData StudyEventOID=\"" + StudyEventOID + "\">");
              out.println("\t\t\t\t<FormData FormOID=\"" + FormOID + "\" OpenClinica:Status=\"initial data entry\">");
              out.println("\t\t\t\t\t\t<ItemGroupData ItemGroupOID=\"IG_CUEST_" + Igroup + "\" " + "ItemGroupRepeatKey=\"" + repeatkey + "\" TransactionType=\"Insert\"" + ">");
              
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_FECHA_" + Item + "\" " + "Value= \"" + fecha + "\"/>");
              out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_HORA_" + Item + "\" " + "Value= \"" + time + "\"/>");
              primer = false;
            }
            if (Igroup.equals("CORAZON"))
            {
              if (formatCell(row.getCell(5)).equals("Tengo los pies más hinchados de lo habitual:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_PIES\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Me siento más fatigado o ahogado de lo habitual:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_FATIGA\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("He pasado mala noche por culpa del ahogo:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_NOCHE\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("He tenido que añadir más almohadas para respirar mejor por la noche:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_ALMOHADAS\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("He tenido que dormir sentado por culpa del ahogo:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_SENTADO\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Me he sentido más mareado o débil de lo habitual:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_MAREADO\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("He tenido mas dolor en el pecho de lo habitual:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_PECHO\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("En general, hoy me siento peor que ayer:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_PEOR\" Value= \"" + row.getCell(7) + "\"/>");
              } else {
                System.out.println("Linea no tratada, sujeto " + formatCell(row.getCell(0)) + " linea " + formatCell(row.getCell(5)));
              }
            }
            else if (Igroup.equals("PULMONES")) {
              if (formatCell(row.getCell(5)).equals("Tengo más ahogo que el usual:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_AHOGO\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Mi esputo ha cambiado de color (o es más obscuro):")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_ESPUTO_COLOR\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Mi esputo ha aumentado:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_ESPUTO_MAS\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Tengo sintomas de resfriado:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_RESFRIADO\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Me han aumentado los 'pitos' en el pecho:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_PITOS\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Tengo dolor de garganta:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_GARGANTA\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Me ha aumentado la tos:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_TOS\" Value= \"" + row.getCell(7) + "\"/>");
              } else if (formatCell(row.getCell(5)).equals("Tengo fiebre:")) {
                out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_FIEBRE\" Value= \"" + row.getCell(7) + "\"/>");
              } else {
                System.out.println("Linea no tratada, sujeto " + formatCell(row.getCell(0)) + " linea " + formatCell(row.getCell(5)));
              }
            }
            fecha_anterior = fecha;
            tiempo_anterior = time;
            Igroup_anterior = Igroup;
            score_global = formatCell(row.getCell(1));
          }
        }
        stmt.close();
        con.close();
      }
      catch (Exception e)
      {
        System.out.println(e.getMessage());
      }
      if (Igroup.equals("CORAZON")) {
        out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_SCORE_COR\" Value= \"" + score_global + "\"/>");
      } else if (Igroup.equals("PULMONES")) {
        out.println("\t\t\t\t\t\t\t<ItemData ItemOID=\"I_CUEST_SCORE_PUL\" Value= \"" + score_global + "\"/>");
      }
      out.println("\t\t\t\t\t\t</ItemGroupData>");
      out.println("\t\t\t\t</FormData>");
      out.println("\t\t\t</StudyEventData>");
      out.println("\t\t</SubjectData>");
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
    formatter.printHelp("cuestionario.jar", options);
  }
}
