package org.hoshisoft.tools.gs;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.hoshisoft.tools.gs.enums.GoogleDocumentsProjection;
import org.hoshisoft.tools.gs.enums.GoogleDocumentsVisibility;

import org.hoshisoft.tools.gs.formatters.*;

import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.PlainTextConstruct;
import com.google.gdata.data.spreadsheet.CustomElementCollection;
import com.google.gdata.data.spreadsheet.ListEntry;
import com.google.gdata.data.spreadsheet.ListFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;

public class GoogleSpreadsheet {
    public static final String SPREADSHEET_FEED_URL =
            "https://spreadsheets.google.com/feeds/spreadsheets";

    public static final String WORKSHEET_FEED_URL =
            "https://spreadsheets.google.com/feeds/worksheets";

    public static final String CELLS_FEED_URL =
            "https://spreadsheets.google.com/feeds/cells";

    public static final String ROWS_FEED_ULR =
            "https://spreadsheets.google.com/feeds/list";

    private String username;
    private String password;
    private String key;
    private String title;
    private String visibility;
    private String projection;

    private Map<String, List<ListEntry>> data = null;

    public Set<String> getLoadedWorksheetTitles() {
        return data.keySet();
    }
    
    /**
     * Creates new instance of class GoogleSpreadsheet.
     * 
     * @param key
     *            Google Spreadsheet key
     * @param title
     *            Google Spreadsheet title
     * @param username
     *            user credentials
     * @param password
     *            user credentials
     * @param visibility
     *            Google Documents visibility
     * @param projection
     *            Google Documents projection
     */
    public GoogleSpreadsheet(
            String key,
            String title,
            String username,
            String password,
            GoogleDocumentsVisibility visibility,
            GoogleDocumentsProjection projection) {

        this.key = key;
        this.title = title;
        this.username = username;
        this.password = password;
        this.visibility = visibility.value();
        this.projection = projection.value();

        this.data = new HashMap<String, List<ListEntry>>();
    }
    
    /**
     * 
     * @param key
     * @param title
     * @param username
     * @param password
     */
    public GoogleSpreadsheet(
            String key,
            String title,
            String username,
            String password) {

        this(
            key,
            title,
            username,
            password,
            GoogleDocumentsVisibility.PRIVATE,
            GoogleDocumentsProjection.FULL);
    }
    

    /**
     * Adds new worksheet with spacified name (title) and dimensions.
     * 
     * @param workSheetName
     *            title of the worksheet
     * @param colCount
     *            number of initial columns
     * @param rowCount
     *            number of initial rows
     * 
     * @throws AuthenticationException
     * @throws MalformedURLException
     * @throws IOException
     * @throws ServiceException
     */
    public void addWorksheet(String workSheetName, int colCount, int rowCount)
            throws AuthenticationException, MalformedURLException, IOException,
            ServiceException {

        // Create and initialize service.
        SpreadsheetService service = initializeService();

        // Define the URL to request.
        URL spreadSheetURL =
                createSpreadsheetURL(
                    SPREADSHEET_FEED_URL,
                    null, // No key needed
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all spreadsheets.
        SpreadsheetFeed feed =
                service.getFeed(spreadSheetURL, SpreadsheetFeed.class);

        List<SpreadsheetEntry> spreadsheets = feed.getEntries();
        if (spreadsheets.size() > 0) {
            // Choose a spreadsheet based on a title.
            for (SpreadsheetEntry spreadsheet : spreadsheets) {
                if (!StringUtils.equals(spreadsheet.getTitle().getPlainText(), this.title)) {
                    continue;
                }

                // Create a local representation of the new worksheet.
                WorksheetEntry worksheet = new WorksheetEntry(colCount, rowCount);
                worksheet.setTitle(new PlainTextConstruct(workSheetName));
                
                // Send the local representation of the worksheet to the API for
                // creation. The URL to use here is the worksheet feed URL of
                // our spreadsheet.
                URL worksheetFeedUrl = spreadsheet.getWorksheetFeedUrl();
                service.insert(worksheetFeedUrl, worksheet);
            }
        }
    }

    
    /**
     * Loads worksheet with specific title.
     * 
     * @param worksheetTitle
     *            worksheet title to be loaded
     * 
     * @throws AuthenticationException
     * @throws MalformedURLException
     * @throws IOException
     * @throws ServiceException
     */
    public void loadWorksheet(String worksheetTitle)
            throws AuthenticationException,
            MalformedURLException,
            IOException,
            ServiceException {

        // Check if we have already loaded entries for input worksheet title.
        if (data.get(worksheetTitle) != null) {
            return;
        }

        // Create and initialize service.
        SpreadsheetService service = initializeService();

        // Define the URL to request.
        URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            List<WorksheetEntry> worksheets = feed.getEntries();

            if (worksheets.size() > 0) {
                for (WorksheetEntry worksheet : worksheets) {
                    // Loop and find the one matching input title.
                    if (worksheet.getTitle().getPlainText().equals(worksheetTitle)) {
                        URL listFeedUrl = worksheet.getListFeedUrl();
                        ListFeed listFeed =
                                service.getFeed(listFeedUrl, ListFeed.class);

                        if (listFeed != null) {
                            this.data.put(worksheetTitle, listFeed.getEntries());
                        }
                    }
                }
            }
        }
    }

    
    /**
     * Loads all worksheets of this spreadsheet.
     * 
     * @throws AuthenticationException
     * @throws MalformedURLException
     * @throws IOException
     * @throws ServiceException
     */
    public void loadAllWorksheets()
            throws AuthenticationException,
            MalformedURLException,
            IOException,
            ServiceException {

        // Create and initialize service.
        SpreadsheetService service = initializeService();

        // Define the URL to request.
        URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            List<WorksheetEntry> worksheets = feed.getEntries();

            if (worksheets.size() > 0) {
                // Loop and load all worksheets.
                for (WorksheetEntry worksheet : worksheets) {
                    URL listFeedUrl = worksheet.getListFeedUrl();
                    ListFeed listFeed =
                            service.getFeed(listFeedUrl, ListFeed.class);

                    if (listFeed != null) {
                        this.data.put(
                            worksheet.getTitle().getPlainText(),
                            listFeed.getEntries());
                    }
                }
            }
        }
    }
    
    
    /**
     * Deletes worksheet specified with worksheetTitle.
     * 
     * @param worksheetTitle title of a worksheet to be deleted.
     * 
     * @throws AuthenticationException
     * @throws MalformedURLException
     * @throws IOException
     * @throws ServiceException
     */
    public void deleteWorksheet(String worksheetTitle)
            throws AuthenticationException,
            MalformedURLException,
            IOException,
            ServiceException {

        // Check if we have already loaded entries for input worksheet title.
        if (worksheetTitle == null) {
            return;
        }
        
        // First, remove local representation of worskheet data.
        data.put(worksheetTitle, null);

        // Create and initialize service.
        SpreadsheetService service = initializeService();

        // Define the URL to request.
        URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            List<WorksheetEntry> worksheets = feed.getEntries();

            if (worksheets.size() > 0) {
                // Loop and find the one matching input title.
                for (WorksheetEntry worksheet : worksheets) {
                    if (worksheet.getTitle().getPlainText().equals(worksheetTitle)) {
                        worksheet.delete();
                    }
                }
            }
        }
    }

    
    /**
     * 
     * @param worksheetTitle
     * @param entries
     * @throws IOException
     * @throws ServiceException
     */
    public void writeToWorksheet(String worksheetTitle, List<ListEntry> entries) 
            throws IOException, ServiceException {
        
        // Check if we have already loaded entries for input worksheet title.
        if (entries == null) {
            return;
        }

        // Create and initialize service.
        SpreadsheetService service = initializeService();

        // Define the URL to request.
        URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    GoogleSpreadsheet.WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            List<WorksheetEntry> worksheets = feed.getEntries();
    
            if (worksheets.size() > 0) {    
                URL listFeedUrl = null;
                
                for (WorksheetEntry worksheet : worksheets) {
                    if (worksheet.getTitle().getPlainText().equals(worksheetTitle)) {
                        listFeedUrl = worksheet.getListFeedUrl();
                        break;
                    }
                }
                
                if (listFeedUrl != null) {
                    for (ListEntry entry : entries) {
                        service.insert(listFeedUrl, entry);
                        
                        try {
                            Thread.sleep(1);
                        } catch (InterruptedException e) {
                            // OK... do nothing...
                        }
                    }
                }
            }
        }
    }
    
    
    /**
     * Prints content of a worksheet, specified with worksheetTitle to stdout.
     * 
     * @param worksheetTitle
     *            the title of worksheet whos content is to be printed
     */
    public void printWorksheet(String worksheetTitle) {
        printWorksheet(worksheetTitle, System.out, new TabGSOutputFormatter());
    }

    
    /**
     * Writes content of a worksheet, specified with worksheetTitle to file.
     * 
     * @param worksheetTitle
     * @param file
     * @throws FileNotFoundException
     */
    public void printWorksheet(String worksheetTitle, File file)
            throws FileNotFoundException {
        printWorksheet(worksheetTitle, new PrintStream(file), new TabGSOutputFormatter());
    }

    
    /**
     * 
     * @param worksheetTitle
     * @param where
     */
    public void printWorksheet(
            String worksheetTitle, 
            PrintStream where, 
            GSOutputFormatter formatter) {
        
        printWorksheet(this.getEntries(worksheetTitle), where, formatter);
    }

    
    /**
     * 
     * @param entries
     * @param where
     */
    protected void printWorksheet(
            List<ListEntry> entries, 
            PrintStream where,
            GSOutputFormatter formatter) {
        
        if (entries != null) {
            // Print header.
            where.println(formatter.format(entries.get(0), true));

            // Print actual data.
            for (ListEntry entry : entries) {
                where.println(formatter.format(entry, false));
            }
        }
    }

    public static final Character TAB = new Character('\t');

    /**
     * Returns tab sepparated string representation of input entry.
     * 
     * @param entry
     * @param isHeader
     * @return
     */
    protected static String listEntryToTSVLine(ListEntry entry, boolean isHeader) {
        StringBuilder sb = new StringBuilder();

        CustomElementCollection elements = entry.getCustomElements();

        boolean isFirst = true;
        for (String tag : elements.getTags()) {
            if (!isFirst) {
                sb.append(TAB);
            } else {
                isFirst = false;
            }

            if (isHeader) {
                sb.append(tag);
            } else {
                sb.append(elements.getValue(tag));
            }
        }

        return sb.toString();
    }

    
    /**
     * Returns list of ListEntry objects for a input worksheetTitle.
     * 
     * @param worksheetTitle
     * @return
     */
    public List<ListEntry> getEntries(String worksheetTitle) {
        return data.get(worksheetTitle);
    }

    public String getUsername() {
        return username;
    }

    public String getPassword() {
        return password;
    }

    public String getKey() {
        return key;
    }

    public String getVisibility() {
        return visibility;
    }

    public String getProjection() {
        return projection;
    }

    
    /**
     * Returns google spreadesheet api v3 URL.
     * 
     * @param scope
     * @param key
     * @param visibility
     * @param projection
     * @return
     * @throws MalformedURLException
     */
    protected static final URL createSpreadsheetURL(
            String scope,
            String key,
            String visibility,
            String projection)

            throws MalformedURLException {
        StringBuilder url = new StringBuilder();

        url.append(scope).append("/");

        if (key != null) {
            url.append(key).append("/");
        }

        url.append(visibility).append("/");
        url.append(projection);

        return new URL(url.toString());
    }

    
    /**
     * Returns object of class SpreadsheetService with authorization rules set. 
     * Default is username / password authorization. Should be overriden in subclasses.
     * 
     * @param service 
     * @return object of class SpreadsheetService with authorization rules set/
     * @throws AuthenticationException
     */
    protected SpreadsheetService authorize(SpreadsheetService service)
            throws AuthenticationException {

        // Setting user credentials.
        if (this.username != null && this.password != null) {
            service.setUserCredentials(this.username, this.password);
        }

        return service;
    }
    
    
    /**
     * Returns newwly created object of class SpreadsheetService with version 
     * and autorization set.
     * 
     * @return newwly created object of class SpreadsheetService and autorization set.
     * @throws AuthenticationException
     */
    private SpreadsheetService initializeService() 
            throws AuthenticationException {
        
        SpreadsheetService service =
                new SpreadsheetService(this.getClass().getName());

        // Setting protocol version to newest (version 3).
        service.setProtocolVersion(SpreadsheetService.Versions.V3);

        // Setting stuff needed for authorization.
        service = this.authorize(service);
        
        return service;
    }
    
//    public void printAllWorksheets(File dir) throws FileNotFoundException {
//        for (Map.Entry<String, List<ListEntry>> entries : data.entrySet()) {
//            // If we have any entries, let's print them to the file.
//            if (entries.getValue() != null) {
//                printWorksheet(
//                    entries.getKey(),
//                    new File(dir.getAbsolutePath() + File.pathSeparator + this.title + "." + entries.getKey()));
//            }
//        }
//    }

}