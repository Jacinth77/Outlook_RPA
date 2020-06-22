package com.novayre.jidoka.robot.test;

import java.io.DataOutputStream;
import java.io.File;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Stream;

import com.novayre.jidoka.outlook.api.model.*;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;

import com.novayre.jidoka.client.api.ECredentialSearch;
import com.novayre.jidoka.client.api.IJidokaRobot;
import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.exceptions.JidokaException;
import com.novayre.jidoka.client.api.exceptions.JidokaFatalException;
import com.novayre.jidoka.client.api.execution.IUsernamePassword;
import com.novayre.jidoka.client.lowcode.IRobotVariable;
import com.novayre.jidoka.outlook.api.IJidokaOutlook;
import com.novayre.jidoka.outlook.api.exception.JidokaMsOutlookException;
import com.novayre.jidoka.windows.api.IWindows;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

/**
 * My robot
 * @author jidoka
 *
 */
@Robot
public class MSOutlookRobot implements IRobot {

	/**
	 * Parameter name for the source Outlook folder.
	 */
	private static final String OUTLOOK_SOURCE_FOLDER_PARAMETER = "OULOOK_SOURCE_FOLDER";

	/**
	 * Parameter name for the target Outlook folder.
	 */
	private static final String OUTLOOK_FOLDER_TARGET_PARAMETER = "OULOOK_FOLDER_TARGET";

	/**
	 * Parameter name for the attachments regular expression use to save them.
	 */
	private static final String ATTACHMENT_REGEXP_PARAMETER = "ATTACHMENT_REGEXP";

	/** The Constant CREDENTIALS_APPIAN_APIKEY. */
	private static final String CREDENTIALS_APPIAN_APIKEY = "GoogleDocs";

	/**
	 * Server.
	 */
	private IJidokaServer< ? > server;

	/**
	 * Windows handler.
	 */
	private IWindows windows;

	/**
	 * Outlook driver.
	 */
	private IJidokaOutlook outlook;

	/**
	 * Email Account.
	 */
	private String emailAccountParam;

	/**
	 * Source folder parameter value.
	 */
	private String folderSourceParam;

	/**
	 * Target folder parameter value.
	 */
	private String folderTargetParam;

	/**
	 * Source & Target folder value.
	 */
	private String sourceFolder;
	private String targetFolder;

	/**
	 * Directory where the attachments will be saved.
	 */
	private File attachmentsDir;

	/**
	 * Mail list.
	 */
	private List<IOlMailItem> mailList;

	/**
	 * Item index.
	 */
	private int currentItemIndex = 1;

	/**
	 * Current Email Index.
	 */
	private int currentEmailIndex = 1;

	/**
	 * Current item.
	 * <p>
	 * Items are the e-mail identifier.
	 */
	private IOlMailItem currentItem;

	private Map<String,IRobotVariable> lowCodeVariables;

	/** The Appian Service Account APIKey. */
	private IUsernamePassword appianCredentials;

	private String uploadUrl;
	private String apiKey;
	/**
	 * Action "start"
	 * @return
	 * @throws Exception
	 */
	public void start() throws JidokaException {

		server = (IJidokaServer<?>) JidokaFactory.getServer();

		windows = IJidokaRobot.getInstance(this);
		outlook = IJidokaOutlook.getInstance(this);

		try {
			appianCredentials = server.getCredential(CREDENTIALS_APPIAN_APIKEY, true, ECredentialSearch.FIRST_LISTED);
			apiKey = appianCredentials.getPassword();
		} catch(Exception e){
			throw new JidokaFatalException("Credentials not found");
		}

		initFromParameters();

		/*
		 * Set the number of items as the total number of e-mails,
		 */
		server.setNumberOfItems(mailList.size());
	}

	public String emails() {
		return currentEmailIndex <= mailList.size() ? "yes" : "no";
	}

	/**
	 * Action 'More Mails?'.
	 * <p>
	 * If there are more e-mails to process.
	 *
	 * @return the output wire name for the action
	 */
	public String moreEmails() {

		// Increase indices
		currentEmailIndex++;
		currentItemIndex++;

		return currentEmailIndex <= mailList.size() ? "yes" : "no";
	}

	/**
	 * Action 'Select Mail'.
	 * <p>
	 * Lists the current mail identifier.
	 */
	public void selectMail() {

		currentItem = mailList.get(currentEmailIndex - 1);

		server.setCurrentItem(currentItemIndex, currentItem.getEntryID());

		server.info(String.format("The current e-mail id is %s and subject %s",
				currentItem.getEntryID(), currentItem.getSubject()));

		setLowCodeVariables();

	}

	/**
	 * Action 'Save Attachments'.
	 * <p>
	 * Saves the attachments of an e-mail.
	 */
	public void saveAttachments() {

		try {

			server.info(String.format(
					"Trying to save the attachments of the e-mail with subject %s",
					currentItem.getSubject()));

			saveMailAttachments(currentItem);

		} catch (JidokaException e) {

			String warnMsg = "Error saving the e-mail attachments";
			server.warn(warnMsg, e);
			server.setCurrentItemResultToWarn(warnMsg);
		}
	}
public void readEmail() {

	List<IOlMailItem> mailList = outlook.getOlFolderManager().getMailList("00000000BFC82839EBA7B3438CC18675EBE64F0582820000");

	for (IOlMailItem mailItem : mailList) {


		server.info("Mail Subject <" + mailItem.getSubject() + ">");
		server.info("     Sender <" + mailItem.getSenderEmailAddress() + ">");
		server.info("     To <" + mailItem.getTo() + ">");
		server.info("     CC <" + mailItem.getCc() + ">");
		server.info("     BCC <" + mailItem.getBcc() + ">");
		server.info("     Body <" + mailItem.getBody() + ">");
		server.info("     CreationTime <" + mailItem.getCreationTime() + ">");
		if(mailItem.getBody() == "Loan Amount 200"){
			break;
		}


	}
}


	public void readEmails() {
		boolean maxvalue = true;

		IOlFolderFW criteria = new OlFolderFW();
		criteria.setFolderPath(sourceFolder);
		List<IOlFolder> folderList = outlook.getOlFolderManager().findFolder(criteria);

		mailList = outlook.getOlFolderManager().getMailList(folderList.get(0).getEntryID());

		for (int i=0;i<mailList.size();i++) {


			server.info("Mail Subject <" + mailList.get(i).getSubject() + ">");
			server.info("     Sender <" + mailList.get(i).getSenderEmailAddress() + ">");
			server.info("     To <" + mailList.get(i).getTo() + ">");
			server.info("     CC <" + mailList.get(i).getCc() + ">");
			server.info("     BCC <" + mailList.get(i).getBcc() + ">");
			server.info("     Body <" + mailList.get(i).getBody() + ">");
			server.info("     CreationTime <" + mailList.get(i).getCreationTime() + ">");
			Document doc = Jsoup.parse(mailList.get(i).getBody());
			String text = doc.body().text();
			int Amt = Integer.parseInt(text.split("Amount")[1]);
			for (int j=0;j<mailList.size();j++){
				Document doce = Jsoup.parse(mailList.get(i).getBody());
				String textvalue = doce.body().text();
				int AmtValue = Integer.parseInt(textvalue.split("Amount")[1]);
				if (Amt == AmtValue){
					j = mailList.size()+1;
					maxvalue = false;

				}

			}

			if(maxvalue){
				i = mailList.size()+1;
				currentItem = mailList.get(currentEmailIndex - 1);

			}
			maxvalue = true;


		}
	}




	/**
	 * Action 'Move Mail'.
	 * <p>
	 * Moves the current e-mail to another folder.
	 */
	public void moveMail() {

		try {
			server.info(String.format("Trying to move the e-mail with subject %s to the folder %s",
					currentItem.getSubject(), targetFolder));
			currentItem.setUnRead(false);
			// Moves it
			IOlFolderFW criteria = new OlFolderFW();
			criteria.setFolderPath(targetFolder);
			IOlFolder target = getFolder(criteria);
			outlook.getOlMailManager().move(currentItem, target);
			server.info(String.format("The e-mail %s has been moved to the folder %s",
					currentItem.getSubject(), targetFolder));

			server.setCurrentItemResultToOK();

			IRobotVariable pId = lowCodeVariables.get("processId");
			server.info("Process Id: " + pId.getValue());

		} catch (JidokaMsOutlookException e) {

			String warnMsg = String.format(
					"The e-mail with subject %s could not be moved", currentItem.getSubject());
			server.warn(warnMsg, e);
			server.setCurrentItemResultToWarn(warnMsg);
		}
	}

	private IOlFolder getFolder(IOlFolderFW criteria) {

		List<IOlFolder> folderList = outlook.getOlFolderManager().findMailFolder(criteria);

		if(folderList.size() > 1) {
			throw new JidokaFatalException("More than 1 folder with the same name");
		}

		IOlFolder target = folderList.get(0);

		return target;
	}

	/**
	 * Action "end"
	 * @return
	 * @throws Exception
	 */
public void end() throws Exception {
		windows.pause(10000);
		outlook.close();
	}


	@Override
	public String[] cleanUp() throws Exception {
		return new String[0];
	}

	private void saveMailAttachments(IOlMailItem mail) throws JidokaException {

		String absolutePath = attachmentsDir.getAbsolutePath();

		outlook.getOlMailManager().downloadAttachments(
				mail, absolutePath);

		server.info(String.format("Attachments of mail with id %s and subject \"%s\" were saved to folder %s",
				mail.getEntryID(), mail.getSubject(), absolutePath));
	}

	public void uploadAttachmentsToAppian() throws JidokaFatalException, IOException {

		server.debug("Looking for files in: " + attachmentsDir.getAbsolutePath());

		List<String> documentIds = new ArrayList();

		try (Stream<Path> files = Files.list(Paths.get(attachmentsDir.getAbsolutePath()))) {
			long count = files.count();
			server.setNumberOfItems((int)count);
		}

		for (final File fileEntry : Objects.requireNonNull(attachmentsDir.listFiles())) {
			server.debug("Found: " + fileEntry.getName());

			String result = uploadFile(fileEntry);

			Map<String, Object> resultMap = (Map<String, Object>) RestHelper.fromJson(result);
			server.debug("result: " + result);

			if ((int)resultMap.get("documentId") > 0) {
				documentIds.add(resultMap.get("documentId").toString());
			} else {
				throw new JidokaFatalException("Issue starting the Appian process.");
			}

		}

		setLowCodeDocVariable(documentIds);
		try{
			if (attachmentsDir.exists()) {
				FileUtils.forceDelete(attachmentsDir);
			}

			FileUtils.forceMkdir(attachmentsDir);
		}
		catch (IOException e){
			server.warn(String.format("%s could not be cleaned", attachmentsDir.getAbsolutePath()));
		}
	}

	private void initFromParameters() throws JidokaFatalException {

		// Set account and folders
		emailAccountParam = server.getParameters().get("emailAccount");
		folderSourceParam = server.getParameters().get("folderSource");
		folderTargetParam = server.getParameters().get("folderTarget");

		// create folder paths
		sourceFolder = String.format("\\\\%s\\%s", emailAccountParam, folderSourceParam);
		targetFolder = String.format("\\\\%s\\%s", emailAccountParam, folderTargetParam);

		// Local directory
		attachmentsDir = new File(server.getCurrentDir() + "\\attachments\\");

		// Cleaning the directory by deleting the previous attachment files
		server.info("Deleting previous downloaded files...");

		try {

			if (attachmentsDir.exists()) {
				FileUtils.forceDelete(attachmentsDir);
			}

			FileUtils.forceMkdir(attachmentsDir);

		} catch (IOException e) {
			server.warn(String.format("%s could not be cleaned"), e);
		}

		uploadUrl = server.getEnvironmentVariables().get("uploadUrl");

		lowCodeVariables = server.getWorkflowVariables();

		server.info("wfv:" + lowCodeVariables.toString());

		for (String wlVariable: lowCodeVariables.keySet()
		) {
			server.info("Workflow:  " + wlVariable);
		}

		//Get all the e-mails IDs inside the source folder
		server.info("Searching email account for source folder: " + sourceFolder);
		IOlFolderFW criteria = new OlFolderFW();
		criteria.setFolderPath(sourceFolder);
		List<IOlFolder> folderList = outlook.getOlFolderManager().findFolder(criteria);

		for (int i = 0; i < folderList.size(); i++) {
			server.info("Folder name  :" + folderList.get(i).getName());
		}

		if(folderList.size() > 1) {
			throw new JidokaFatalException("More than 1 folder with the same name");
		}

		if(folderList.isEmpty()){
			throw new JidokaFatalException("Source Folder Not Found!");
		}

		mailList = outlook.getOlFolderManager().getMailList(folderList.get(0).getEntryID());
	}


	private String uploadFile(File file) throws JidokaFatalException, IOException {

		HttpURLConnection conn = null;

		String result;
		try {
			URL url = new URL(uploadUrl);
			conn = (HttpURLConnection)url.openConnection();
			conn.setRequestProperty("Appian-API-Key", apiKey);
			conn.setRequestProperty("Content-Disposition", "attachment; filename=" + file.getName());
			conn.setRequestProperty("Content-Type", "application/octet-stream");
			conn.setRequestProperty("Appian-Document-Name", file.getName());
			conn.setDoOutput(true);
			conn.setRequestMethod("POST");
			DataOutputStream wr = new DataOutputStream(conn.getOutputStream());
			wr.write(FileUtils.readFileToByteArray(file));
			wr.flush();
			wr.close();
			if (conn.getResponseCode() != 200) {
				throw new JidokaFatalException("Web API request failed:" + conn.getResponseCode());
			}

			result = IOUtils.toString(conn.getInputStream());
		}
		catch (IOException e){
			throw new JidokaFatalException("Timeout Exception, check the Environment Variable 'uploadUrl' in the Robot Configuration");
		}
		finally {
			if (conn != null) {
				conn.disconnect();
			}
		}

		return result;
	}

	private void setLowCodeDocVariable(List<String> documentIds){
		lowCodeVariables.get("documentIds").setValue(documentIds.toArray());
	}

	private void setLowCodeVariables () {
		lowCodeVariables.get("emailBody").setValue(currentItem.getBody());
		lowCodeVariables.get("fromEmail").setValue(currentItem.getSenderEmailAddress());
		lowCodeVariables.get("fromName").setValue(currentItem.getSenderName());
		lowCodeVariables.get("subject").setValue(currentItem.getSubject());
	}
}

