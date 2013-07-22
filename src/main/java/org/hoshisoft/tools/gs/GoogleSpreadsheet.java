/**
 *   Copyright 2013 Luka Obradovic
 *
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */

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

    private Map<String, List<ListEntry>> data;

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
            final String key,
            final String title,
            final String username,
            final String password,
            final GoogleDocumentsVisibility visibility,
            final GoogleDocumentsProjection projection) {

        this.key = key;
        this.title = title;
        this.username = username;
        this.password = password;
        this.visibility = visibility.value();
        this.projection = projection.value();

        this.data = new HashMap<String, List<ListEntry>>();
    }
    
    /**
     * Creates new instance of class GoogleSpreadsheet.
     * 
     * @param key
     * @param title
     * @param username
     * @param password
     */
    public GoogleSpreadsheet(
            final String key,
            final String title,
            final String username,
            final String password) {

        this(
            key,
            title,
            username,
            password,
            GoogleDocumentsVisibility.PRIVATE,
            GoogleDocumentsProjection.FULL);
    }
    

    /**
     * Adds new worksheet spacified by <code>workSheetName</code> and dimensions.
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
    public void addWorksheet(
            final String workSheetName,
            final int colCount,
            final int rowCount)

    throws  AuthenticationException,
            MalformedURLException,
            IOException,
            ServiceException {

        // Create and initialize service.
        final SpreadsheetService service = initializeService();

        // Define the URL to request.
        final URL spreadSheetURL =
                createSpreadsheetURL(
                    SPREADSHEET_FEED_URL,
                    null, // No key needed
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all spreadsheets.
        final SpreadsheetFeed feed =
                service.getFeed(spreadSheetURL, SpreadsheetFeed.class);

        final List<SpreadsheetEntry> spreadsheets = feed.getEntries();
        if (!spreadsheets.isEmpty()) {
            // Choose a spreadsheet based on a title.
            for (SpreadsheetEntry spreadsheet : spreadsheets) {
                if (!StringUtils.equals(spreadsheet.getTitle().getPlainText(), this.title)) {
                    continue;
                }

                // Create a local representation of the new worksheet.
                final WorksheetEntry worksheet = new WorksheetEntry(colCount, rowCount);
                worksheet.setTitle(new PlainTextConstruct(workSheetName));
                
                // Send the local representation of the worksheet to the API for
                // creation. The URL to use here is the worksheet feed URL of
                // our spreadsheet.
                final URL worksheetFeedUrl = spreadsheet.getWorksheetFeedUrl();
                service.insert(worksheetFeedUrl, worksheet);
            }
        }
    }

    
    /**
     * Loads worksheet specified with <code>worksheetTitle</code>.
     * 
     * @param worksheetTitle
     *            worksheet title to be loaded
     * 
     * @throws AuthenticationException
     * @throws MalformedURLException
     * @throws IOException
     * @throws ServiceException
     */
    public void loadWorksheet(final String worksheetTitle)
            throws AuthenticationException,
            MalformedURLException,
            IOException,
            ServiceException {

        // Check if we have already loaded entries for input worksheet title.
        if (data.get(worksheetTitle) != null) {
            return;
        }

        // Create and initialize service.
        final SpreadsheetService service = initializeService();

        // Define the URL to request.
        final URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        final WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            final List<WorksheetEntry> worksheets = feed.getEntries();

            if (worksheets.size() > 0) {
                for (WorksheetEntry worksheet : worksheets) {
                    // Loop and find the one matching input title.
                    if (worksheet.getTitle().getPlainText().equals(worksheetTitle)) {
                        final URL listFeedUrl = worksheet.getListFeedUrl();

                        final ListFeed listFeed =
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
        final SpreadsheetService service = initializeService();

        // Define the URL to request.
        final URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        final WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            final List<WorksheetEntry> worksheets = feed.getEntries();

            if (!worksheets.isEmpty()) {
                // Loop and load all worksheets.
                for (WorksheetEntry worksheet : worksheets) {
                    final URL listFeedUrl = worksheet.getListFeedUrl();

                    final ListFeed listFeed =
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
     * Deletes worksheet specified with <code>worksheetTitle</code>.
     * 
     * @param worksheetTitle title of a worksheet to be deleted.
     * 
     * @throws AuthenticationException
     * @throws MalformedURLException
     * @throws IOException
     * @throws ServiceException
     */
    public void deleteWorksheet(final String worksheetTitle)
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
        final SpreadsheetService service = initializeService();

        // Define the URL to request.
        final URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        final WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            final List<WorksheetEntry> worksheets = feed.getEntries();

            if (!worksheets.isEmpty()) {
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
    public void writeToWorksheet(
            final String worksheetTitle,
            final List<ListEntry> entries)
    throws IOException, ServiceException {
        
        // Check if we have already loaded entries for input worksheet title.
        if (entries == null) {
            return;
        }

        // Create and initialize service.
        final SpreadsheetService service = initializeService();

        // Define the URL to request.
        final URL URL_FEED_REQUEST =
                createSpreadsheetURL(
                    GoogleSpreadsheet.WORKSHEET_FEED_URL,
                    this.key,
                    this.visibility,
                    this.projection);

        // Make a request to the API and get all worksheets.
        final WorksheetFeed feed =
                service.getFeed(URL_FEED_REQUEST, WorksheetFeed.class);

        if (feed != null) {
            final List<WorksheetEntry> worksheets = feed.getEntries();
    
            if (!worksheets.isEmpty()) {
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
     * Prints content of worksheet, specified with <code>worksheetTitle</code> 
     * to a standard out in tsv format.
     * 
     * @param worksheetTitle
     *            the title of worksheet whos content is to be printed
     */
    public void printWorksheet(final String worksheetTitle) {
        printWorksheet(worksheetTitle, System.out, new TabGSOutputFormatter());
    }

    
    /**
     * Prints content of worksheet, specified with <code>worksheetTitle</code> 
     * to a <code>file</code> in tsv format.
     * 
     * @param worksheetTitle
     * @param file
     * @throws FileNotFoundException
     */
    public void printWorksheet(final String worksheetTitle, final File file)
            throws FileNotFoundException {
        printWorksheet(
                worksheetTitle,
                new PrintStream(file),
                new TabGSOutputFormatter());
    }

    
    /**
     * Prints content of worksheet, specified with <code>worksheetTitle</code> 
     * to <code>where</code>, formating output with <code>formatter</code>.
     * 
     * @param worksheetTitle
     * @param where
     */
    public void printWorksheet(
            final String worksheetTitle,
            final PrintStream where,
            final GSOutputFormatter formatter) {
        
        printWorksheet(this.getEntries(worksheetTitle), where, formatter);
    }

    
    /**
     * Prints <code>entries</code> to <code>where</code>, formatting output with
     * <code>formatter</code>.
     *
     * @param entries
     * @param where
     */
    protected void printWorksheet(
            final List<ListEntry> entries,
            final PrintStream where,
            final GSOutputFormatter formatter) {
        
        if (entries != null) {
            // Print header.
            where.println(formatter.format(entries.get(0), true));

            // Print actual data.
            for (ListEntry entry : entries) {
                where.println(formatter.format(entry, false));
            }
        }
    }
    

    /**
     * Returns list of ListEntry objects for a input worksheetTitle.
     * 
     * @param worksheetTitle
     * @return
     */
    public List<ListEntry> getEntries(final String worksheetTitle) {
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

    public static final char URL_PATH_SEPARATOR = '/';
    
    /**
     * Returns google spreadsheet api v3 URL.
     * 
     * @param scope
     * @param key
     * @param visibility
     * @param projection
     * @return
     * @throws MalformedURLException
     */
    protected static final URL createSpreadsheetURL(
            final String scope,
            final String key,
            final String visibility,
            final String projection)
    throws MalformedURLException {

        final StringBuilder url = new StringBuilder();

        url.append(scope).append(URL_PATH_SEPARATOR);

        if (key != null) {
            url.append(key).append(URL_PATH_SEPARATOR);
        }

        url.append(visibility).append(URL_PATH_SEPARATOR);
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
    protected void authorize(final SpreadsheetService service)
            throws AuthenticationException {

        // Setting user credentials.
        if (this.username != null && this.password != null) {
            service.setUserCredentials(this.username, this.password);
        }
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

        final SpreadsheetService service =
                new SpreadsheetService(this.getClass().getName());

        // Setting protocol version to newest (version 3).
        service.setProtocolVersion(SpreadsheetService.Versions.V3);

        // Setting stuff needed for authorization.
        this.authorize(service);
        
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