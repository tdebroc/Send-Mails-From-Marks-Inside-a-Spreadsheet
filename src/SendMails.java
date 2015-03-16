import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.channels.Channels;
import java.nio.channels.ReadableByteChannel;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import au.com.bytecode.opencsv.CSV;
import au.com.bytecode.opencsv.CSVReadProc;

public class SendMails {
	//=========================================================================
	//= Config
	//=========================================================================
	/* Key for the Google Spreadsheet containing marks.  */
	final static String googleSpreadhseetKey = "1JobjeymNWivpwTH1x2syL9pUill9kilF8Bx1ro19lnA";
	/* Index of the line for Questions title. */
	final static int QUESTION = 0;
	/* Index of the line for bonus value. */
	final static int BONUS = 2;
	/* Index of the line indicating max mark for a question. */
	final static int MAX_MARK = 3;
	/* Index of the first line of student. */
	final static int FIRST_STUDENT = 4;
	//=========================================================================
	//= Global attributes for the class.
	//=========================================================================
	static ArrayList<ArrayList<String>> allCsv;
	static Transport t;
	static Session session;
	static Properties mailProps;
	
	
	
	//=========================================================================
	//= Methods.
	//=========================================================================
	/**
	 * Main Code.
	 */
	public static void main(String[] args) throws MessagingException {
		initSendMail();
		readMarksAndSendMails();
	}

	/**
	 * Reads marks from the Google Docs, create mails and send them.
	 */
	public static void readMarksAndSendMails() {
		CSV csv = CSV.create();

		URL website;
		try {
			website = new URL(
					"https://docs.google.com/spreadsheets/d/" + googleSpreadhseetKey + "/export?format=csv");
			ReadableByteChannel rbc = Channels.newChannel(website.openStream());
			FileOutputStream fos = new FileOutputStream("marks.csv");
			fos.getChannel().transferFrom(rbc, 0, Long.MAX_VALUE);
			fos.close();
		} catch (MalformedURLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		allCsv = new ArrayList<ArrayList<String>>();
		csv.read("marks.csv", new CSVReadProc() {
			public void procRow(int rowIndex, String... values) {
				allCsv.add(new ArrayList<String>(Arrays.asList(values)));
			}
		});

		ArrayList<String> questions = allCsv.get(QUESTION);
		ArrayList<String> bonuses = allCsv.get(BONUS);
		ArrayList<String> maxMarks = allCsv.get(MAX_MARK);

		for (int i = FIRST_STUDENT; !allCsv.get(i).get(0).equals(""); i++) {
			ArrayList<String> student = allCsv.get(i);
			StringBuilder mail = new StringBuilder();
			mail.append("Bonjour " + student.get(0) + ",<br/><br/>");
			mail.append("Voilà le compte rendu de votre note de TP: <br/>");

			int j = 1;
			for (j = 1; !questions.get(j).equals("Note"); j++) {
				String mark = student.get(j);
				mail.append("<u>" + questions.get(j));
				mail.append(!bonuses.get(j).equals("0") ? " (bonus)" : "");
				mail.append("</u>: " + mark + "/" + maxMarks.get(j) + "<br/>");
			}
			mail.append("<b>Total: " + student.get(j) + " / 20 ("
					+ maxMarks.get(j) + " avec bonus) </b>");

			mail.append("<br/><br/>");
			mail.append("Voilà des Commentaires personnel sur votre TP:<br/>");
			mail.append(student.get(j + 3));

			if (!student.get(j + 4).equals("")) {
				mail.append("<br/><br/>");
				mail.append("Au TP2, J'avais précisé que je ne soushaitais pas voir "
						+ "du code copié/collé et souhaitais que l'on m'informe "
						+ "lorsque le TP était fait en binôme.<br/>");
				mail.append("Malheureusement j'ai retrouvé de gros copier/coller avec "
						+ "plusieurs autres TPs, d'où une très légère pénalité sur "
						+ "ce TP en espérant grandement que cela ne se reproduise plus.<br/>");
			}

			mail.append("<br/><br/>");

			mail.append("Vous pourrez retrouver des commentaires, des conseils "
					+ "et des corrigés sur le lien suivant:<br/>"
					+ "<a href=\"https://docs.google.com/document/d/1DxFSQ6kBFaHTR1A7CINJ_oKaEw_N87f1_KVNa1CI6g8\">"
					+ "https://docs.google.com/document/d/1DxFSQ6kBFaHTR1A7CINJ_oKaEw_N87f1_KVNa1CI6g8"
					+ "</a>");
			mail.append("<br/><br/>");
			mail.append("Bien cordialement,<br/><br/>");
			mail.append("Thibaut de Broca");

			String mails = student.get(j + 2);
			System.out.println("Send mail to " + mails);
			sendMail(mails, "TP2 Java - Feedback", mail.toString());

			System.out.println(mail.toString().replaceAll("<br/>", "\n")
					+ "\n\n\n");

		}

	}

	/**
	 * Sends Mail to persons.
	 * 
	 * @param toMails
	 *            {String} Mails separated by a ";".
	 * @param subject
	 *            Subject of mail.
	 * @param text
	 *            Content of mail.
	 */
	public static void sendMail(String toMails, String subject, String text) {
		MimeMessage message = new MimeMessage(session);
		String from = mailProps.getProperty("username");
		try {
			message.setFrom(new InternetAddress(from));
			for (String mail : toMails.split(";")) {
				message.addRecipient(Message.RecipientType.TO,
						new InternetAddress(mail));
			}
			message.setSubject(subject);
			message.setText(text, "utf-8", "html");

			t.sendMessage(message, message.getAllRecipients());
		} catch (AddressException e) {
			e.printStackTrace();
		} catch (MessagingException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Inits Messaging System Connection.
	 * 
	 * @throws MessagingException
	 */
	public static void initSendMail() throws MessagingException {
		loadMailProperties();
		Properties props = new Properties();
		Session session = Session.getInstance(props);
		t = session.getTransport("smtps");
		t.connect(mailProps.getProperty("host"),
				mailProps.getProperty("username"),
				mailProps.getProperty("password"));
	}

	/**
	 * Gets Properties needed for sending mail.
	 * 
	 * @return
	 */
	public static void loadMailProperties() {
		Properties prop = new Properties();
		String propFileName = "resources/mailConfig.properties";
		FileInputStream fis;
		try {
			fis = new FileInputStream(propFileName);
			try {
				if (fis != null) {
					prop.load(fis);
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (FileNotFoundException e1) {
			System.err.println("Please Provide a " + propFileName + " file.");
			e1.printStackTrace();
		}

	}
}
