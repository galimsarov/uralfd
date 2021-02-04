import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.*;

public class CheckingMails {
    private static final Scanner scanner = new Scanner(System.in);
    private static String username;
    private static String password;
    private static String recipientAddress;

    public static void main
            (String[] args) throws InterruptedException {
        String host = "pop.gmail.com";
        String pathName = "/home/egalimsarov/Загрузки/Файл отказы.xlsx";
        System.out.println("Добрый день! Укажите, пожалуйста, email, " +
                "входящую почту на котором будем обрабатывать:");
        username = scanner.nextLine();
        System.out.println("Для подключеня к рабочему почтовому ящику "
                + username + " необходимо ввести пароль:");
        password = scanner.nextLine();
        System.out.println("Укажите, пожалуйста, адрес получателя результатов " +
                "обработки:");
        recipientAddress = scanner.nextLine();
        while (true) {
            check(host, username, password, pathName);
            System.out.println("---------");
            System.out.println("Почтовый ящик проверен");
            Calendar calendar = Calendar.getInstance();
            System.out.println("Время: " + calendar.get(Calendar.HOUR_OF_DAY));
            if (calendar.get(Calendar.HOUR_OF_DAY) == 18) {
                System.out.println("Время 12:00, завершаем работу!");
                break;
            }
            else {
                System.out.println("Продолжаем работу");
                Thread.sleep(5 * 60 * 1000);
            }
        }
    }

    public static void check
            (String host, String user, String password, String pathName) {
        try {
            //create properties field
            Properties properties = new Properties();

            properties.put("mail.pop3.host", host);
            properties.put("mail.pop3.port", "995");
            properties.put("mail.pop3.starttls.enable", "true");
            Session emailSession = Session.getDefaultInstance(properties);

            //create the POP3 store object and connect with the pop server
            Store store = emailSession.getStore("pop3s");

            store.connect(host, user, password);

            //create the folder object and open it
            Folder emailFolder = store.getFolder("INBOX");
            emailFolder.open(Folder.READ_ONLY);

            // retrieve the messages from the folder in an array and print it
            Message[] messages = emailFolder.getMessages();
            System.out.println("Количество входящих писем: " + messages.length);

            for (Message message : messages) {
                if (message.getSubject().contains("заполнена web-форма")) {
                    Object msgContent = message.getContent();
                    String content = "";
                    // Check if content is pure text/html or in parts
                    if (msgContent instanceof Multipart) {
                        Multipart multipart = (Multipart) msgContent;
                        for (int j = 0; j < multipart.getCount(); j++) {
                            BodyPart bodyPart = multipart.getBodyPart(j);
                            content = bodyPart.getContent().toString();
                        }
                    } else
                        content = msgContent.toString();
                    String[] parts = content.split("Телефон");
                    String phoneNumber = parts[1];
                    parts = phoneNumber.split("Дата рождения");
                    phoneNumber = parts[0].substring(33)
                            .replace("\n", "")
                            .replace("\r", "")
                            .replace("+", "")
                            .replace("(", "")
                            .replace(")", "")
                            .replace(" ", "")
                            .replace("-", "");
                    String birthDay = parts[1].substring(33, 45)
                            .replace("\n", "")
                            .replace("\r", "");
                    String[] dateAndNumber = new String[2];
                    dateAndNumber[0] = birthDay;
                    dateAndNumber[1] = phoneNumber;
                    find(pathName, dateAndNumber, message);
                }
            }

            //close the store and folder objects
            emailFolder.close(false);
            store.close();
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    public static void find
            (String pathName, String[] dateAndNumber, Message message) {
        try {
            System.out.println("---------");
            System.out.println("Ищем клиента с датой рождения " + dateAndNumber[0] +
                    " и телефоном " + dateAndNumber[1]);
            //creating a new file instance
            File file = new File(pathName);
            //obtaining bytes from the file
            FileInputStream fis = new FileInputStream(file);
            //creating Workbook instance that refers to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            //creating a Sheet object to retrieve object
            XSSFSheet sheet = wb.getSheetAt(0);
            //iterating over excel file
            Iterator<Row> itr = sheet.iterator();
            boolean emailSent = false;
            while (itr.hasNext()) {
                Row row = itr.next();
                //iterating over each column
                Iterator<Cell> cellIterator = row.cellIterator();
                boolean birthDayEquals = false, clientFound = false;
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType() == CellType.STRING)
                        if (cell.getStringCellValue().equals(dateAndNumber[0])) {
                            // дата рождения совпала
                            // System.out.println("Совпала дата рождения: " + dateAndNumber[0]);
                            birthDayEquals = true;
                        }
                    if (cell.getCellType() == CellType.NUMERIC) {
                        double phoneFromFile = cell.getNumericCellValue();
                        String phoneNumber = BigDecimal.valueOf(phoneFromFile).toPlainString();
                        if (birthDayEquals && phoneNumber.substring(1).
                                equals(dateAndNumber[1].substring(1)))
                            clientFound = true;
                    }
                    if (clientFound)
                        break;
                }
                if (clientFound) {
                    emailSent = true;
                    sendMail(message, true);
                    break;
                }
            }
            if (!emailSent)
                sendMail(message, false);
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void sendMail
            (Message message, boolean clientFound) {
        Properties prop = new Properties();
        prop.put("mail.smtp.host", "smtp.gmail.com");
        prop.put("mail.smtp.port", "587");
        prop.put("mail.smtp.auth", "true");
        prop.put("mail.smtp.starttls.enable", "true"); //TLS

        Session session = Session.getInstance(prop,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });

        try {
            MimeMessage newMessage = new MimeMessage(session);
            String from = message.getRecipients(Message.RecipientType.TO)[0].toString()
                    .split("<")[1]
                    .split(">")[0];
            System.out.println("Письмо от: " + from);
            String to = message.getFrom()[0].toString()
                    .split("<")[1]
                    .split(">")[0];
            System.out.println("Письмо кому: " + to);
            System.out.println("Клиент новый: " + !clientFound);

            newMessage.setFrom(new InternetAddress(from));
            newMessage.addRecipient(Message.RecipientType.TO,
                    new InternetAddress(recipientAddress));
            if (!clientFound)
                newMessage.setSubject("NEW CLIENT! " + message.getSubject());
            else
                newMessage.setSubject("BAD CLIENT! " + message.getSubject());
            newMessage.setText(message.getContent().toString());

            Transport.send(newMessage);
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}
