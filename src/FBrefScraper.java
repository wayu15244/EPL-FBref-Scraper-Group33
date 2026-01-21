import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.ZonedDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Random;

/**
 * FBref Premier League Match Scraper
 * 
 * Scrapes ALL 380 matches from the 2024-2025 Premier League season from
 * FBref.com.
 * Extracts: Teams, Score, Date, Time (UK), Stadium, Referee, Attendance,
 * Goal Scorers, Starting XI, Possession, Shots, Shots on Target
 * 
 * Uses Selenium WebDriver for JavaScript-heavy pages.
 * Uses Apache POI for Excel output.
 * 
 * @author Group Task 33
 * @version 1.0
 */
public class FBrefScraper {

    // FBref URLs
    private static final String FIXTURES_URL = "https://fbref.com/en/comps/9/2024-2025/schedule/2024-2025-Premier-League-Scores-and-Fixtures";
    private static final String BASE_URL = "https://fbref.com";
    private static final String OUTPUT_FILE = "EPL_2024_2025_FBref.xlsx";

    // UK Time Zone
    private static final ZoneId UK_ZONE = ZoneId.of("Europe/London");

    // Delay between requests (milliseconds)
    private static final int REQUEST_DELAY = 4000;

    // Selenium driver
    private WebDriver driver;
    private JavascriptExecutor js;
    private WebDriverWait wait;

    // Test mode limit
    private int testLimit = -1; // -1 = no limit

    public static void main(String[] args) {
        System.out.println("============================================================");
        System.out.println("  FBref Premier League Scraper");
        System.out.println("  Season 2024-2025 (380 Matches)");
        System.out.println("  Group Task 33");
        System.out.println("============================================================");
        System.out.println();

        FBrefScraper scraper = new FBrefScraper();

        // Check for test mode
        for (String arg : args) {
            if (arg.equals("--test")) {
                scraper.testLimit = 5;
                System.out.println("[MODE] Test mode enabled - scraping only 5 matches");
            } else if (arg.startsWith("--test=")) {
                scraper.testLimit = Integer.parseInt(arg.substring(7));
                System.out.println("[MODE] Test mode enabled - scraping only " + scraper.testLimit + " matches");
            }
        }

        scraper.run();
    }

    public void run() {
        try {
            System.out.println("[Step 1] Starting Chrome browser...");
            initializeDriver();

            System.out.println("[Step 2] Loading FBref fixtures page...");
            driver.get(FIXTURES_URL);
            sleep(5000);

            System.out.println("[Step 3] Collecting match URLs...");
            List<String> matchUrls = collectMatchUrls();

            if (testLimit > 0 && matchUrls.size() > testLimit) {
                matchUrls = matchUrls.subList(0, testLimit);
            }

            System.out.println("  Found " + matchUrls.size() + " matches to scrape");

            if (matchUrls.isEmpty()) {
                System.out.println("ERROR: No match URLs found!");
                return;
            }

            System.out.println("[Step 4] Extracting match details...");
            int estimatedMinutes = matchUrls.size() * (REQUEST_DELAY / 1000) / 60;
            System.out.println("  (Estimated time: " + estimatedMinutes + " minutes)");

            List<MatchData> allMatches = new ArrayList<>();
            Random random = new Random();

            for (int i = 0; i < matchUrls.size(); i++) {
                // Progress update
                if ((i + 1) % 10 == 0 || i == 0) {
                    System.out.println("  Processing " + (i + 1) + "/" + matchUrls.size() + "...");
                }

                try {
                    // Random delay to simulate human behavior (3-6 seconds)
                    if (i > 0) {
                        int randomDelay = 3000 + random.nextInt(3000);
                        sleep(randomDelay);
                    }

                    MatchData match = scrapeMatch(matchUrls.get(i));
                    if (match != null && !match.homeTeam.isEmpty()) {
                        allMatches.add(match);

                        // Print first match as sample
                        if (i == 0) {
                            System.out.println("    First match: " + match.homeTeam + " vs " + match.awayTeam + " ("
                                    + match.fullTimeScore + ")");
                        }
                    }
                } catch (Exception e) {
                    System.out.println("    Error on match " + (i + 1) + ": " + e.getMessage());
                }

                // Base delay between requests
                sleep(REQUEST_DELAY);
            }

            System.out.println("[Step 5] Writing to Excel file...");
            writeToExcel(allMatches);

            System.out.println();
            System.out.println("============================================================");
            System.out.println("  SUCCESS!");
            System.out.println("  Total matches scraped: " + allMatches.size());
            System.out.println("  Output file: " + OUTPUT_FILE);
            System.out.println("============================================================");

        } catch (Exception e) {
            System.out.println("ERROR: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
            }
        }
    }

    /**
     * Initialize Chrome WebDriver with STEALTH options to avoid anti-bot detection
     * Uses a custom Chrome profile for the scraper
     */
    private void initializeDriver() {
        ChromeOptions options = new ChromeOptions();

        // === USE CUSTOM CHROME PROFILE for scraper ===
        // This creates a persistent profile that can store cookies
        String userHome = System.getProperty("user.home");
        String scraperProfilePath = userHome + "\\AppData\\Local\\FBrefScraper\\ChromeProfile";
        options.addArguments("--user-data-dir=" + scraperProfilePath);

        // === STEALTH MODE: Hide automation signals ===
        // Remove automation flags
        options.setExperimentalOption("excludeSwitches", Arrays.asList("enable-automation"));
        options.setExperimentalOption("useAutomationExtension", false);
        options.addArguments("--disable-blink-features=AutomationControlled");

        // Disable automation-related flags
        options.addArguments("--disable-infobars");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--no-sandbox");
        options.addArguments("--remote-debugging-port=9222");

        // Basic options
        options.addArguments("--start-maximized");
        options.addArguments("--disable-notifications");
        options.addArguments("--disable-popup-blocking");
        options.addArguments("--lang=en-GB");

        // Disable webdriver flag
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("credentials_enable_service", false);
        prefs.put("profile.password_manager_enabled", false);
        options.setExperimentalOption("prefs", prefs);

        driver = new ChromeDriver(options);

        // Hide webdriver property via JavaScript
        js = (JavascriptExecutor) driver;
        try {
            js.executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})");
        } catch (Exception e) {
            // Ignore if this fails
        }

        wait = new WebDriverWait(driver, Duration.ofSeconds(20));
    }

    /**
     * Collect all match report URLs from the fixtures page
     */
    private List<String> collectMatchUrls() {
        List<String> urls = new ArrayList<>();

        try {
            // Wait for table to load
            wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("table")));

            // Get all match report links
            // FBref uses td[data-stat="match_report"] for match report links
            List<WebElement> matchLinks = driver.findElements(By.cssSelector("td[data-stat=\"match_report\"] a"));

            if (matchLinks.isEmpty()) {
                // Fallback: try score links
                matchLinks = driver.findElements(By.cssSelector("td[data-stat=\"score\"] a"));
            }

            for (WebElement link : matchLinks) {
                String href = link.getAttribute("href");
                if (href != null && href.contains("/matches/")) {
                    if (!href.startsWith("http")) {
                        href = BASE_URL + href;
                    }
                    if (!urls.contains(href)) {
                        urls.add(href);
                    }
                }
            }

            System.out.println("  Found " + urls.size() + " match report links");

        } catch (Exception e) {
            System.out.println("  Error collecting URLs: " + e.getMessage());
        }

        return urls;
    }

    /**
     * Scrape data from a single match page
     */
    private MatchData scrapeMatch(String matchUrl) {
        MatchData match = new MatchData();

        try {
            driver.get(matchUrl);
            sleep(2000);

            // Wait for scorebox to load
            wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("div.scorebox")));

            // Extract all data
            extractTeamsAndScore(match);
            extractMatchInfo(match);
            extractGoalScorers(match);
            extractManagers(match);
            extractLineups(match);
            extractSubstitutes(match);
            extractStats(match);
            extractExtraStats(match);
            extractHTScoreAndCards(match);

            // Calculate HT score from goal minutes (more reliable than DOM scraping)
            calculateHTScore(match);

            // FALLBACK: Global search for missing advanced stats
            globalKeywordSearch(match);

        } catch (Exception e) {
            // Return partial data if available
        }

        return match;
    }

    /**
     * Extract team names and final score
     */
    private void extractTeamsAndScore(MatchData match) {
        try {
            // Home team
            WebElement homeTeamEl = driver.findElement(By.cssSelector("div.scorebox > div:first-child strong a"));
            match.homeTeam = homeTeamEl.getText().trim();

            // Away team
            WebElement awayTeamEl = driver.findElement(By.cssSelector("div.scorebox > div:nth-child(2) strong a"));
            match.awayTeam = awayTeamEl.getText().trim();

            // Scores
            List<WebElement> scores = driver.findElements(By.cssSelector("div.scorebox div.score"));
            if (scores.size() >= 2) {
                String homeScore = scores.get(0).getText().trim();
                String awayScore = scores.get(1).getText().trim();
                match.fullTimeScore = homeScore + " - " + awayScore;
            }

        } catch (Exception e) {
            // Try alternative selectors
            try {
                String script = "var scorebox = document.querySelector('div.scorebox'); " +
                        "if (!scorebox) return null; " +
                        "var teams = scorebox.querySelectorAll('strong a'); " +
                        "var scores = scorebox.querySelectorAll('div.score'); " +
                        "return { " +
                        "  home: teams[0] ? teams[0].innerText : '', " +
                        "  away: teams[1] ? teams[1].innerText : '', " +
                        "  homeScore: scores[0] ? scores[0].innerText : '', " +
                        "  awayScore: scores[1] ? scores[1].innerText : '' " +
                        "};";

                Object result = js.executeScript(script);
                if (result instanceof java.util.Map) {
                    java.util.Map<?, ?> data = (java.util.Map<?, ?>) result;
                    match.homeTeam = getStr(data, "home");
                    match.awayTeam = getStr(data, "away");
                    match.fullTimeScore = getStr(data, "homeScore") + " - " + getStr(data, "awayScore");
                }
            } catch (Exception ex) {
            }
        }
    }

    /**
     * Extract match info: Date, Time, Stadium, Attendance, Referee
     */
    private void extractMatchInfo(MatchData match) {
        try {
            // FBref stores match info in div.scorebox_meta
            driver.findElement(By.cssSelector("div.scorebox_meta"));

            // Parse date and time (UK venue time)
            // Format example: "Friday August 16, 2024" and "20:00 (20:00 venue time)"
            String script = "var meta = document.querySelector('div.scorebox_meta'); " +
                    "if (!meta) return {}; " +
                    "var text = meta.innerText; " +
                    "var result = {}; " +

                    // Date
                    "var dateMatch = text.match(/([A-Za-z]+\\s+[A-Za-z]+\\s+\\d{1,2},\\s+\\d{4})/); " +
                    "if (dateMatch) result.date = dateMatch[1]; " +

                    // Venue time (UK time)
                    "var timeMatch = text.match(/(\\d{1,2}:\\d{2})\\s*\\(.*?venue time\\)/i); " +
                    "if (timeMatch) result.time = timeMatch[1]; " +
                    "if (!timeMatch) { " +
                    "  timeMatch = text.match(/Start Time:\\s*(\\d{1,2}:\\d{2})/i); " +
                    "  if (timeMatch) result.time = timeMatch[1]; " +
                    "} " +

                    // Venue/Stadium
                    "var venueMatch = text.match(/Venue:\\s*([^\\n]+)/i); " +
                    "if (venueMatch) result.venue = venueMatch[1].trim(); " +

                    // Attendance
                    "var attMatch = text.match(/Attendance:\\s*([\\d,]+)/i); " +
                    "if (attMatch) result.attendance = attMatch[1]; " +

                    // Referee - find span containing (Referee) and extract name
                    "var spans = meta.querySelectorAll('span'); " +
                    "for (var i = 0; i < spans.length; i++) { " +
                    "  if (spans[i].textContent.includes('(Referee)')) { " +
                    "    result.referee = spans[i].textContent.replace('(Referee)', '').trim().replace(/\\u00A0/g, ' '); "
                    +
                    "    break; " +
                    "  } " +
                    "} " +
                    // Fallback regex if span not found
                    "if (!result.referee) { " +
                    "  var refMatch = text.match(/Officials:[^·]*?([A-Za-z\\s]+)\\s*\\(Referee\\)/i); " +
                    "  if (refMatch) result.referee = refMatch[1].trim(); " +
                    "} " +

                    "return result;";

            Object result = js.executeScript(script);
            if (result instanceof java.util.Map) {
                java.util.Map<?, ?> info = (java.util.Map<?, ?>) result;

                String dateStr = getStr(info, "date");
                String timeStr = getStr(info, "time");

                // Format as UK time
                match.dateTime = formatUKDateTime(dateStr, timeStr);
                match.stadium = getStr(info, "venue");
                match.attendance = getStr(info, "attendance");
                match.referee = getStr(info, "referee");
            }

        } catch (Exception e) {
        }
    }

    /**
     * Format date and time as UK local time
     */
    private String formatUKDateTime(String dateStr, String timeStr) {
        if (dateStr == null || dateStr.isEmpty())
            return "";

        try {
            // Parse date: "Friday August 16, 2024"
            DateTimeFormatter inputFmt = DateTimeFormatter.ofPattern("EEEE MMMM d, yyyy", Locale.ENGLISH);
            LocalDate date = LocalDate.parse(dateStr, inputFmt);

            // Parse time if available
            String timeFormatted = "";
            if (timeStr != null && !timeStr.isEmpty()) {
                timeFormatted = " " + timeStr;
            }

            // Format output
            DateTimeFormatter outputFmt = DateTimeFormatter.ofPattern("dd MMM yyyy", Locale.ENGLISH);
            ZonedDateTime ukTime = date.atStartOfDay(UK_ZONE);

            // Determine if BST or GMT
            String zoneAbbr = ukTime.getZone().getRules().isDaylightSavings(ukTime.toInstant()) ? "BST" : "GMT";

            return date.format(outputFmt) + timeFormatted + " " + zoneAbbr;

        } catch (Exception e) {
            return dateStr + (timeStr != null ? " " + timeStr : "");
        }
    }

    /**
     * Extract goal scorers with minutes
     */
    private void extractGoalScorers(MatchData match) {
        try {
            String script = "var scorers = []; " +
                    "var events = document.querySelectorAll('div.event'); " +
                    "events.forEach(function(e) { " +
                    "  var text = e.innerText; " +
                    "  if (text.includes('Goal') || text.match(/\\d+'/)) { " +
                    "    scorers.push(text.replace(/\\n/g, ' ').trim()); " +
                    "  } " +
                    "}); " +
                    // Fallback: get from scorebox
                    "if (scorers.length === 0) { " +
                    "  document.querySelectorAll('div.scorebox div[class*=\"event\"]').forEach(function(e) { " +
                    "    scorers.push(e.innerText.replace(/\\n/g, ' ').trim()); " +
                    "  }); " +
                    "} " +
                    "return scorers.join('; ');";

            js.executeScript(script);
            // Goal scorers info is now stored in homeGoals/awayGoals via
            // extractHTScoreAndCards

        } catch (Exception e) {
        }
    }

    /**
     * Extract starting XI for both teams
     */
    private void extractLineups(MatchData match) {
        try {
            String script = "var result = { homeXI: [], awayXI: [] }; " +

            // Method 1: Try div.lineup
                    "var lineups = document.querySelectorAll('div.lineup'); " +
                    "if (lineups.length >= 2) { " +
                    "  lineups[0].querySelectorAll('a').forEach(function(a, i) { " +
                    "    if (i < 11) result.homeXI.push(a.innerText.trim()); " +
                    "  }); " +
                    "  lineups[1].querySelectorAll('a').forEach(function(a, i) { " +
                    "    if (i < 11) result.awayXI.push(a.innerText.trim()); " +
                    "  }); " +
                    "} " +

                    // Method 2: Try player tables
                    "if (result.homeXI.length === 0) { " +
                    "  var tables = document.querySelectorAll('table[id*=\"stats\"][id*=\"summary\"]'); " +
                    "  if (tables.length >= 2) { " +
                    "    tables[0].querySelectorAll('tbody tr').forEach(function(tr, i) { " +
                    "      if (i < 11) { " +
                    "        var name = tr.querySelector('th a'); " +
                    "        if (name) result.homeXI.push(name.innerText.trim()); " +
                    "      } " +
                    "    }); " +
                    "    tables[1].querySelectorAll('tbody tr').forEach(function(tr, i) { " +
                    "      if (i < 11) { " +
                    "        var name = tr.querySelector('th a'); " +
                    "        if (name) result.awayXI.push(name.innerText.trim()); " +
                    "      } " +
                    "    }); " +
                    "  } " +
                    "} " +

                    "return result;";

            Object result = js.executeScript(script);
            if (result instanceof java.util.Map) {
                java.util.Map<?, ?> data = (java.util.Map<?, ?>) result;

                Object homeList = data.get("homeXI");
                Object awayList = data.get("awayXI");

                if (homeList instanceof List) {
                    match.homeStartingXI = String.join(", ",
                            ((List<?>) homeList).stream().map(Object::toString).toArray(String[]::new));
                }
                if (awayList instanceof List) {
                    match.awayStartingXI = String.join(", ",
                            ((List<?>) awayList).stream().map(Object::toString).toArray(String[]::new));
                }
            }

        } catch (Exception e) {
        }
    }

    /**
     * Extract match statistics: Possession, Shots, Shots on Target
     * THIS IS CRITICAL - must not be N/A
     * 
     * FBref team_stats structure:
     * - Row 1: Team names
     * - Row 2: "Possession" header
     * - Row 3: Possession values (55% / 45%) in <strong> tags
     * - Row 4: "Passing Accuracy" header
     * - Row 5: Passing values
     * - Row 6: "Shots on Target" header
     * - Row 7: Shot values "5 of 14 — 36%" / "20% — 2 of 10"
     */
    private void extractStats(MatchData match) {
        try {
            String script = "var stats = { " +
                    "  homePoss: '', awayPoss: '', " +
                    "  homeShots: '', awayShots: '', " +
                    "  homeSOT: '', awaySOT: '' " +
                    "}; " +

                    "var teamStats = document.getElementById('team_stats'); " +
                    "if (teamStats) { " +
                    "  var table = teamStats.querySelector('table'); " +
                    "  if (table) { " +
                    "    var rows = table.querySelectorAll('tbody tr'); " +

                    // Possession is in row 3 (index 2), in <strong> tags
                    "    if (rows.length > 2) { " +
                    "      var possRow = rows[2]; " +
                    "      var strongs = possRow.querySelectorAll('strong'); " +
                    "      if (strongs.length >= 2) { " +
                    "        stats.homePoss = strongs[0].innerText.trim(); " +
                    "        stats.awayPoss = strongs[1].innerText.trim(); " +
                    "      } " +
                    "    } " +

                    // Shots on Target is in row 7 (index 6)
                    "    if (rows.length > 6) { " +
                    "      var shotRow = rows[6]; " +
                    "      var cells = shotRow.querySelectorAll('td'); " +
                    "      if (cells.length >= 2) { " +
                    "        var homeText = cells[0].innerText; " +
                    "        var awayText = cells[1].innerText; " +
                    // Parse "5 of 14" format - could be "5 of 14 — 36%" or "20% — 2 of 10"
                    "        var homeMatch = homeText.match(/(\\d+)\\s*of\\s*(\\d+)/i); " +
                    "        var awayMatch = awayText.match(/(\\d+)\\s*of\\s*(\\d+)/i); " +
                    "        if (homeMatch) { " +
                    "          stats.homeSOT = homeMatch[1]; " +
                    "          stats.homeShots = homeMatch[2]; " +
                    "        } " +
                    "        if (awayMatch) { " +
                    "          stats.awaySOT = awayMatch[1]; " +
                    "          stats.awayShots = awayMatch[2]; " +
                    "        } " +
                    "      } " +
                    "    } " +
                    "  } " +
                    "} " +

                    // Fallback: Try alternative selectors if above fails
                    "if (!stats.homePoss) { " +
                    "  var text = document.body.innerText; " +
                    "  var possMatch = text.match(/(\\d+)%\\s*Possession\\s*(\\d+)%/i); " +
                    "  if (possMatch) { " +
                    "    stats.homePoss = possMatch[1] + '%'; " +
                    "    stats.awayPoss = possMatch[2] + '%'; " +
                    "  } else { " +
                    "    possMatch = text.match(/Possession[\\s\\S]{0,50}(\\d+)%[\\s\\S]{0,30}(\\d+)%/i); " +
                    "    if (possMatch) { " +
                    "      stats.homePoss = possMatch[1] + '%'; " +
                    "      stats.awayPoss = possMatch[2] + '%'; " +
                    "    } " +
                    "  } " +
                    "} " +

                    // Debug: log what we found
                    "console.log('Stats extracted:', stats); " +

                    "return stats;";

            Object result = js.executeScript(script);
            if (result instanceof java.util.Map) {
                java.util.Map<?, ?> stats = (java.util.Map<?, ?>) result;
                match.homePossession = getStr(stats, "homePoss");
                match.awayPossession = getStr(stats, "awayPoss");
                match.homeTotalShots = getStr(stats, "homeShots");
                match.awayTotalShots = getStr(stats, "awayShots");
                match.homeShotsOnTarget = getStr(stats, "homeSOT");
                match.awayShotsOnTarget = getStr(stats, "awaySOT");
            }

        } catch (Exception e) {
        }
    }

    /**
     * Extract managers for both teams
     */
    private void extractManagers(MatchData match) {
        try {
            String script = "var managers = []; " +
                    "document.querySelectorAll('.scorebox strong').forEach(function(strong) { " +
                    "  if (strong.textContent.includes('Manager')) { " +
                    "    var parent = strong.parentElement; " +
                    "    var link = parent.querySelector('a'); " +
                    "    if (link) managers.push(link.textContent.trim()); " +
                    "    else managers.push(parent.textContent.replace('Manager:', '').trim()); " +
                    "  } " +
                    "}); " +
                    "return managers;";

            Object result = js.executeScript(script);
            if (result instanceof List) {
                List<?> managers = (List<?>) result;
                if (managers.size() >= 2) {
                    match.homeManager = managers.get(0).toString();
                    match.awayManager = managers.get(1).toString();
                } else if (managers.size() == 1) {
                    match.homeManager = managers.get(0).toString();
                }
            }
        } catch (Exception e) {
        }
    }

    /**
     * Extract substitute players with substitution minutes
     * Format: "Player Name (Min')" for players who came on
     */
    private void extractSubstitutes(MatchData match) {
        try {
            String script = "var result = { homeSubs: [], awaySubs: [] }; " +

            // Get substitution events with minutes - use class 'a' for home, 'b' for away
                    "var events = Array.from(document.querySelectorAll('#events_wrap .event')); " +

                    "events.forEach(function(event) { " +
                    "  var text = event.innerText; " +
                    "  if (text.includes('for')) { " +
                    // Substitution event
                    "    var timeDiv = event.querySelector('div:first-child'); " +
                    "    var time = timeDiv ? timeDiv.textContent.trim().split('\\n')[0].trim() : ''; " +
                    "    var links = event.querySelectorAll('a'); " +
                    "    var playerIn = links[0] ? links[0].textContent.trim() : ''; " +
                    "    if (playerIn && time) { " +
                    "      var subText = playerIn + ' (' + time + ')'; " +
                    // FBref uses class 'a' for home events, 'b' for away events
                    "      var parent = event.parentElement; " +
                    "      var isHome = event.classList.contains('a') || " +
                    "                   (parent && parent.classList.contains('a')); " +
                    "      var isAway = event.classList.contains('b') || " +
                    "                   (parent && parent.classList.contains('b')); " +
                    // Fallback: check position in events container
                    "      if (!isHome && !isAway) { " +
                    "        var eventContainer = event.closest('#events_wrap') || event.parentElement; " +
                    "        var allEvents = Array.from(eventContainer.querySelectorAll('.event')); " +
                    "        var idx = allEvents.indexOf(event); " +
                    // Check if event is on left or right side based on text alignment
                    "        var style = window.getComputedStyle(event); " +
                    "        isHome = style.textAlign === 'left' || idx % 2 === 0; " +
                    "        isAway = !isHome; " +
                    "      } " +
                    "      if (isHome) result.homeSubs.push(subText); " +
                    "      else result.awaySubs.push(subText); " +
                    "    } " +
                    "  } " +
                    "}); " +

                    // Also get bench players who didn't play (from lineup section)
                    "var benchHeaders = Array.from(document.querySelectorAll('div, th')).filter(function(el) { " +
                    "  return el.textContent.trim() === 'Bench'; " +
                    "}); " +

                    "benchHeaders.forEach(function(header, index) { " +
                    "  var container = header.closest('table') || header.parentElement; " +
                    "  var links = container.querySelectorAll('a'); " +
                    "  links.forEach(function(a) { " +
                    "    if (a.href && a.href.includes('/players/')) { " +
                    "      var rect = a.getBoundingClientRect(); " +
                    "      var headerRect = header.getBoundingClientRect(); " +
                    "      if (rect.top > headerRect.top) { " +
                    "        var playerName = a.textContent.trim(); " +
                    // Check if this player is already in subs list (came on)
                    "        var alreadyIn = (index === 0 ? result.homeSubs : result.awaySubs).some(function(s) { " +
                    "          return s.includes(playerName); " +
                    "        }); " +
                    "        if (!alreadyIn) { " +
                    // Player didn't play - just add name without minutes
                    "          if (index === 0) result.homeSubs.push(playerName); " +
                    "          else result.awaySubs.push(playerName); " +
                    "        } " +
                    "      } " +
                    "    } " +
                    "  }); " +
                    "}); " +

                    "return result;";

            Object result = js.executeScript(script);
            if (result instanceof java.util.Map) {
                java.util.Map<?, ?> data = (java.util.Map<?, ?>) result;

                Object homeList = data.get("homeSubs");
                Object awayList = data.get("awaySubs");

                if (homeList instanceof List) {
                    match.homeSubstitutes = String.join(", ",
                            ((List<?>) homeList).stream().map(Object::toString).toArray(String[]::new));
                }
                if (awayList instanceof List) {
                    match.awaySubstitutes = String.join(", ",
                            ((List<?>) awayList).stream().map(Object::toString).toArray(String[]::new));
                }
            }
        } catch (Exception e) {
        }
    }

    /**
     * Extract Half-Time Score, Goal Events, Cards, and Match Officials
     */
    private void extractHTScoreAndCards(MatchData match) {
        try {
            String script = "var result = { htScore: '', homeGoals: [], awayGoals: [], homeCards: [], awayCards: [], officials: {} }; "
                    +

                    // HT Score from scorebox
                    "var scorebox = document.querySelector('.scorebox'); " +
                    "if (scorebox) { " +
                    "  var text = scorebox.innerText; " +
                    "  var htMatch = text.match(/\\((\\d+[–-]\\d+)\\)/); " +
                    "  if (htMatch) result.htScore = htMatch[1].replace('–', '-'); " +
                    "} " +
                    "if (!result.htScore) result.htScore = '0-0'; " +

                    // Goal and Card events
                    "var events = document.querySelectorAll('#events_wrap .event'); " +
                    "events.forEach(function(e) { " +
                    "  var isHome = e.classList.contains('a') || e.closest('.a') !== null; " +
                    "  var time = e.querySelector('div:first-child')?.textContent.trim().split('\\n')[0].replace(/[^0-9'+]/g, ''); "
                    +

                    // Goals
                    "  var goalIcon = e.querySelector('[class*=\"goal\"]'); " +
                    "  if (goalIcon && !e.innerText.includes('Penalty miss')) { " +
                    "    var links = e.querySelectorAll('a'); " +
                    "    var scorer = links[0]?.textContent.trim() || ''; " +
                    "    var assist = links.length > 1 ? links[1].textContent.trim() : ''; " +
                    "    if (scorer) { " +
                    "      var goalData = scorer + '|' + time + '|' + assist; " +
                    "      if (isHome) result.homeGoals.push(goalData); " +
                    "      else result.awayGoals.push(goalData); " +
                    "    } " +
                    "  } " +

                    // Cards
                    "  var cardIcon = e.querySelector('.yellow_card, .yellow_red_card, .red_card'); " +
                    "  if (cardIcon) { " +
                    "    var player = e.querySelector('a')?.textContent.trim(); " +
                    "    var cardType = cardIcon.classList.contains('red_card') ? 'R' : 'Y'; " +
                    "    if (player) { " +
                    "      var cardData = player + ' ' + time; " +
                    "      if (isHome) result.homeCards.push(cardData); " +
                    "      else result.awayCards.push(cardData); " +
                    "    } " +
                    "  } " +
                    "}); " +

                    // Officials from scorebox_meta (not scorebox - it has too much text)
                    "var meta = document.querySelector('.scorebox_meta'); " +
                    "if (meta) { " +
                    "  var spans = meta.querySelectorAll('span'); " +
                    "  for (var i = 0; i < spans.length; i++) { " +
                    "    var spanText = spans[i].textContent; " +
                    "    if (spanText.includes('(Referee)')) result.officials.referee = spanText.replace('(Referee)', '').trim().replace(/\\u00A0/g, ' '); "
                    +
                    "    if (spanText.includes('(AR1)')) result.officials.ar1 = spanText.replace('(AR1)', '').trim().replace(/\\u00A0/g, ' '); "
                    +
                    "    if (spanText.includes('(AR2)')) result.officials.ar2 = spanText.replace('(AR2)', '').trim().replace(/\\u00A0/g, ' '); "
                    +
                    "    if (spanText.includes('(4th)')) result.officials.fourth = spanText.replace('(4th)', '').trim().replace(/\\u00A0/g, ' '); "
                    +
                    "    if (spanText.includes('(VAR)')) result.officials.var = spanText.replace('(VAR)', '').trim().replace(/\\u00A0/g, ' '); "
                    +
                    "  } " +
                    "} " +

                    "return result;";

            Object result = js.executeScript(script);
            if (result instanceof java.util.Map) {
                java.util.Map<?, ?> data = (java.util.Map<?, ?>) result;
                match.halfTimeScore = getStr(data, "htScore");

                // Goals
                Object homeGoalsList = data.get("homeGoals");
                Object awayGoalsList = data.get("awayGoals");
                if (homeGoalsList instanceof List) {
                    match.homeGoals = String.join(";",
                            ((List<?>) homeGoalsList).stream().map(Object::toString).toArray(String[]::new));
                }
                if (awayGoalsList instanceof List) {
                    match.awayGoals = String.join(";",
                            ((List<?>) awayGoalsList).stream().map(Object::toString).toArray(String[]::new));
                }

                // Cards
                Object homeCardsList = data.get("homeCards");
                Object awayCardsList = data.get("awayCards");
                if (homeCardsList instanceof List) {
                    match.homeCards = String.join(", ",
                            ((List<?>) homeCardsList).stream().map(Object::toString).toArray(String[]::new));
                }
                if (awayCardsList instanceof List) {
                    match.awayCards = String.join(", ",
                            ((List<?>) awayCardsList).stream().map(Object::toString).toArray(String[]::new));
                }

                // Officials
                Object officials = data.get("officials");
                if (officials instanceof java.util.Map) {
                    java.util.Map<?, ?> off = (java.util.Map<?, ?>) officials;
                    if (!getStr(off, "referee").isEmpty())
                        match.referee = getStr(off, "referee");
                    String ar1 = getStr(off, "ar1");
                    String ar2 = getStr(off, "ar2");
                    match.assistantRefs = ar1 + (ar1.isEmpty() || ar2.isEmpty() ? "" : ", ") + ar2;
                    match.fourthOfficial = getStr(off, "fourth");
                    match.varOfficial = getStr(off, "var");
                }
            }
        } catch (Exception e) {
        }
    }

    /**
     * Extract extra stats: xG, Passes, Tackles, Corners, Fouls, Offsides, Cards,
     * Saves, and ALL ADVANCED STATS from #team_stats_extra
     * 
     * FBref structure for #team_stats_extra:
     * Each stat row has 3 divs: [homeValue] [Label] [awayValue]
     */
    private void extractExtraStats(MatchData match) {
        try {
            String script = "var stats = { " +
                    "  homeXG: '', awayXG: '', " +
                    "  homePasses: '', awayPasses: '', " +
                    "  homeTackles: '', awayTackles: '', " +
                    "  homeCorners: '', awayCorners: '', " +
                    "  homeFouls: '', awayFouls: '', " +
                    "  homeOffsides: '', awayOffsides: '', " +
                    "  homeYellows: '', awayYellows: '', " +
                    "  homeReds: '', awayReds: '', " +
                    "  homeSaves: '', awaySaves: '', " +
                    // NEW: Advanced stats for Zone D
                    "  homeBigChances: '0', awayBigChances: '0', " +
                    "  homeWoodwork: '0', awayWoodwork: '0', " +
                    "  homeClearances: '0', awayClearances: '0', " +
                    "  homeInterceptions: '0', awayInterceptions: '0', " +
                    "  homeBlocks: '0', awayBlocks: '0', " +
                    "  homeAerials: '0', awayAerials: '0', " +
                    "  homeCrosses: '0', awayCrosses: '0', " +
                    "  homeLongBalls: '0', awayLongBalls: '0', " +
                    "  homeThroughBalls: '0', awayThroughBalls: '0' " +
                    "}; " +

                    // xG from scorebox using .score_xg class
                    "var xgElements = document.querySelectorAll('.score_xg'); " +
                    "if (xgElements.length >= 2) { " +
                    "  stats.homeXG = xgElements[0].textContent.trim(); " +
                    "  stats.awayXG = xgElements[1].textContent.trim(); " +
                    "} else { " +
                    // Fallback: search in scorebox for xG label pattern
                    "  var scorebox = document.querySelector('.scorebox'); " +
                    "  if (scorebox) { " +
                    "    var xgDivs = Array.from(scorebox.querySelectorAll('div')).filter(function(d) { " +
                    "      return d.textContent.trim() === 'xG'; " +
                    "    }); " +
                    "    if (xgDivs.length > 0) { " +
                    "      var xgContainer = xgDivs[0].parentElement; " +
                    "      var values = Array.from(xgContainer.querySelectorAll('div')).filter(function(d) { " +
                    "        return /^\\d+\\.\\d+$/.test(d.textContent.trim()); " +
                    "      }); " +
                    "      if (values.length >= 2) { " +
                    "        stats.homeXG = values[0].textContent.trim(); " +
                    "        stats.awayXG = values[1].textContent.trim(); " +
                    "      } " +
                    "    } " +
                    "  } " +
                    "} " +

                    // Passes from #team_stats
                    "var teamStats = document.querySelector('#team_stats'); " +
                    "if (teamStats) { " +
                    "  var rows = Array.from(teamStats.querySelectorAll('tr')); " +
                    "  var passingRow = rows.find(function(r) { return r.textContent.includes('Passing Accuracy'); }); "
                    +
                    "  if (passingRow) { " +
                    "    var passingDataRow = passingRow.nextElementSibling; " +
                    "    if (passingDataRow) { " +
                    "      var cells = passingDataRow.querySelectorAll('td'); " +
                    "      if (cells.length >= 2) { " +
                    "        stats.homePasses = cells[0].textContent.trim().split('—')[0].trim(); " +
                    "        stats.awayPasses = cells[1].textContent.trim().split('—')[0].trim(); " +
                    "      } " +
                    "    } " +
                    "  } " +
                    "} " +

                    // ============================================================
                    // COMPREHENSIVE EXTRACTION FROM #team_stats_extra
                    // Each stat row: [HomeValue] [Label] [AwayValue]
                    // ============================================================
                    "var extraStats = document.querySelector('#team_stats_extra'); " +
                    "if (extraStats) { " +
                    "  var allDivs = Array.from(extraStats.querySelectorAll('div')); " +

                    // Helper function to find stat by label text
                    "  function findStat(labelTexts) { " +
                    "    for (var i = 0; i < labelTexts.length; i++) { " +
                    "      var labelDiv = allDivs.find(function(d) { " +
                    "        var text = d.textContent.trim().toLowerCase(); " +
                    "        return text === labelTexts[i].toLowerCase(); " +
                    "      }); " +
                    "      if (labelDiv && labelDiv.previousElementSibling && labelDiv.nextElementSibling) { " +
                    "        return { " +
                    "          home: labelDiv.previousElementSibling.textContent.trim(), " +
                    "          away: labelDiv.nextElementSibling.textContent.trim() " +
                    "        }; " +
                    "      } " +
                    "    } " +
                    "    return { home: '0', away: '0' }; " +
                    "  } " +

                    // Extract Tackles
                    "  var tackles = findStat(['Tackles']); " +
                    "  stats.homeTackles = tackles.home; " +
                    "  stats.awayTackles = tackles.away; " +

                    // Extract Corners
                    "  var corners = findStat(['Corners']); " +
                    "  stats.homeCorners = corners.home; " +
                    "  stats.awayCorners = corners.away; " +

                    // Extract Fouls
                    "  var fouls = findStat(['Fouls']); " +
                    "  stats.homeFouls = fouls.home; " +
                    "  stats.awayFouls = fouls.away; " +

                    // Extract Offsides
                    "  var offsides = findStat(['Offsides']); " +
                    "  stats.homeOffsides = offsides.home; " +
                    "  stats.awayOffsides = offsides.away; " +

                    // ===== ATTACK STATS =====
                    // Big Chances (may be labeled "Big Chances" or "SCA" on FBref)
                    "  var bigChances = findStat(['Big Chances', 'SCA', 'Shot-Creating Actions']); " +
                    "  stats.homeBigChances = bigChances.home; " +
                    "  stats.awayBigChances = bigChances.away; " +

                    // Hit Woodwork / Post
                    "  var woodwork = findStat(['Woodwork', 'Post', 'Hit Woodwork', 'Shots off Woodwork']); " +
                    "  stats.homeWoodwork = woodwork.home; " +
                    "  stats.awayWoodwork = woodwork.away; " +

                    // Crosses
                    "  var crosses = findStat(['Crosses', 'Crs']); " +
                    "  stats.homeCrosses = crosses.home; " +
                    "  stats.awayCrosses = crosses.away; " +

                    // ===== DEFENCE STATS =====
                    // Clearances
                    "  var clearances = findStat(['Clearances', 'Clr']); " +
                    "  stats.homeClearances = clearances.home; " +
                    "  stats.awayClearances = clearances.away; " +

                    // Interceptions
                    "  var interceptions = findStat(['Interceptions', 'Int']); " +
                    "  stats.homeInterceptions = interceptions.home; " +
                    "  stats.awayInterceptions = interceptions.away; " +

                    // Blocks
                    "  var blocks = findStat(['Blocks', 'Blk']); " +
                    "  stats.homeBlocks = blocks.home; " +
                    "  stats.awayBlocks = blocks.away; " +

                    // Aerial Duels Won
                    "  var aerials = findStat(['Aerials Won', 'Aerial Duels', 'Aerial Duels Won', 'Aerials']); " +
                    "  stats.homeAerials = aerials.home; " +
                    "  stats.awayAerials = aerials.away; " +

                    // ===== POSSESSION STATS =====
                    // Long Balls
                    "  var longBalls = findStat(['Long Balls', 'Long', 'Long Passes']); " +
                    "  stats.homeLongBalls = longBalls.home; " +
                    "  stats.awayLongBalls = longBalls.away; " +

                    // Through Balls
                    "  var throughBalls = findStat(['Through Balls', 'Through', 'ThrBalls']); " +
                    "  stats.homeThroughBalls = throughBalls.home; " +
                    "  stats.awayThroughBalls = throughBalls.away; " +
                    "} " +

                    // Yellow cards from #team_stats (count .yellow_card spans)
                    "var teamStats = document.querySelector('#team_stats'); " +
                    "if (teamStats) { " +
                    "  var cardRows = Array.from(teamStats.querySelectorAll('tr')).filter(function(tr) { " +
                    "    return tr.textContent.includes('Cards'); " +
                    "  }); " +
                    "  if (cardRows.length > 0) { " +
                    "    var nextRow = cardRows[0].nextElementSibling; " +
                    "    if (nextRow) { " +
                    "      var cells = nextRow.querySelectorAll('td'); " +
                    "      if (cells.length >= 2) { " +
                    "        stats.homeYellows = cells[0].querySelectorAll('.yellow_card').length.toString(); " +
                    "        stats.awayYellows = cells[1].querySelectorAll('.yellow_card').length.toString(); " +
                    "        stats.homeReds = cells[0].querySelectorAll('.red_card').length.toString(); " +
                    "        stats.awayReds = cells[1].querySelectorAll('.red_card').length.toString(); " +
                    "      } " +
                    "    } " +
                    "  } " +

                    // Saves from Saves row
                    "  var savesRows = Array.from(teamStats.querySelectorAll('tr')).filter(function(tr) { " +
                    "    return tr.textContent.includes('Saves'); " +
                    "  }); " +
                    "  if (savesRows.length > 0) { " +
                    "    var savesDataRow = savesRows[0].nextElementSibling; " +
                    "    if (savesDataRow) { " +
                    "      var cells = savesDataRow.querySelectorAll('td'); " +
                    "      if (cells.length >= 2) { " +
                    // Parse "2 of 2 — 100%" format, extract first number
                    "        var homeMatch = cells[0].textContent.match(/(\\d+)\\s*of/); " +
                    "        var awayMatch = cells[1].textContent.match(/(\\d+)\\s*of/); " +
                    "        if (homeMatch) stats.homeSaves = homeMatch[1]; " +
                    "        if (awayMatch) stats.awaySaves = awayMatch[1]; " +
                    "      } " +
                    "    } " +
                    "  } " +
                    "} " +

                    // Debug output
                    "console.log('ExtraStats extracted:', stats); " +

                    "return stats;";

            Object result = js.executeScript(script);
            if (result instanceof java.util.Map) {
                java.util.Map<?, ?> data = (java.util.Map<?, ?>) result;

                // Advanced Stats
                match.homeXG = getStr(data, "homeXG");
                match.awayXG = getStr(data, "awayXG");
                match.homePasses = getStr(data, "homePasses");
                match.awayPasses = getStr(data, "awayPasses");
                match.homeTackles = getStr(data, "homeTackles");
                match.awayTackles = getStr(data, "awayTackles");

                // Discipline
                match.homeCorners = getStr(data, "homeCorners");
                match.awayCorners = getStr(data, "awayCorners");
                match.homeFouls = getStr(data, "homeFouls");
                match.awayFouls = getStr(data, "awayFouls");
                match.homeOffsides = getStr(data, "homeOffsides");
                match.awayOffsides = getStr(data, "awayOffsides");
                match.homeYellowCards = getStr(data, "homeYellows");
                match.awayYellowCards = getStr(data, "awayYellows");
                match.homeRedCards = getStr(data, "homeReds");
                match.awayRedCards = getStr(data, "awayReds");

                // Defence
                match.homeSaves = getStr(data, "homeSaves");
                match.awaySaves = getStr(data, "awaySaves");
                match.homeClearances = getStrOrDefault(data, "homeClearances", "0");
                match.awayClearances = getStrOrDefault(data, "awayClearances", "0");
                match.homeInterceptions = getStrOrDefault(data, "homeInterceptions", "0");
                match.awayInterceptions = getStrOrDefault(data, "awayInterceptions", "0");
                match.homeBlocks = getStrOrDefault(data, "homeBlocks", "0");
                match.awayBlocks = getStrOrDefault(data, "awayBlocks", "0");
                match.homeAerials = getStrOrDefault(data, "homeAerials", "0");
                match.awayAerials = getStrOrDefault(data, "awayAerials", "0");

                // Attack
                match.homeBigChances = getStrOrDefault(data, "homeBigChances", "0");
                match.awayBigChances = getStrOrDefault(data, "awayBigChances", "0");
                match.homeWoodwork = getStrOrDefault(data, "homeWoodwork", "0");
                match.awayWoodwork = getStrOrDefault(data, "awayWoodwork", "0");
                match.homeCrosses = getStrOrDefault(data, "homeCrosses", "0");
                match.awayCrosses = getStrOrDefault(data, "awayCrosses", "0");

                // Possession
                match.homeLongBalls = getStrOrDefault(data, "homeLongBalls", "0");
                match.awayLongBalls = getStrOrDefault(data, "awayLongBalls", "0");
                match.homeThroughBalls = getStrOrDefault(data, "homeThroughBalls", "0");
                match.awayThroughBalls = getStrOrDefault(data, "awayThroughBalls", "0");
            }
        } catch (Exception e) {
            // Silently handle errors
        }
    }

    /**
     * Helper: Get string from map with default value if empty
     */
    private String getStrOrDefault(java.util.Map<?, ?> map, String key, String defaultVal) {
        Object val = map.get(key);
        if (val == null)
            return defaultVal;
        String str = val.toString().trim();
        return str.isEmpty() ? defaultVal : str;
    }

    /**
     * Helper: Get string from map safely
     */
    private String getStr(java.util.Map<?, ?> map, String key) {
        Object val = map.get(key);
        return val != null ? val.toString().trim() : "";
    }

    /**
     * Calculate HT Score from goal minutes
     * Counts goals scored in first half (minute <= 45 or 45+)
     */
    private void calculateHTScore(MatchData match) {
        try {
            int homeHT = 0;
            int awayHT = 0;

            // Count home goals in first half
            if (!match.homeGoals.isEmpty()) {
                String[] goals = match.homeGoals.split(";");
                for (String goal : goals) {
                    if (goal.isEmpty())
                        continue;
                    String[] parts = goal.split("\\|");
                    if (parts.length > 1) {
                        String minute = parts[1].trim().replaceAll("[^0-9+]", "");
                        if (isFirstHalf(minute)) {
                            homeHT++;
                        }
                    }
                }
            }

            // Count away goals in first half
            if (!match.awayGoals.isEmpty()) {
                String[] goals = match.awayGoals.split(";");
                for (String goal : goals) {
                    if (goal.isEmpty())
                        continue;
                    String[] parts = goal.split("\\|");
                    if (parts.length > 1) {
                        String minute = parts[1].trim().replaceAll("[^0-9+]", "");
                        if (isFirstHalf(minute)) {
                            awayHT++;
                        }
                    }
                }
            }

            match.halfTimeScore = homeHT + "-" + awayHT;
        } catch (Exception e) {
            match.halfTimeScore = "0-0";
        }
    }

    /**
     * Check if minute is in first half (1-45 or 45+)
     */
    private boolean isFirstHalf(String minute) {
        if (minute.isEmpty())
            return false;
        try {
            // Handle "45+2" format
            if (minute.contains("+")) {
                String[] parts = minute.split("\\+");
                int baseMin = Integer.parseInt(parts[0]);
                return baseMin <= 45;
            }
            int min = Integer.parseInt(minute);
            return min <= 45;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * GLOBAL KEYWORD SEARCH: Extract stats from specific FBref tables
     * CORRECTED based on browser inspection:
     * - Through Balls: data-stat="through_balls" (NOT thru_balls)
     * - Big Chances: data-stat="xg_shot" (NOT xg) with value >= 0.35
     * - Woodwork: data-stat="outcome" for "Post" (NOT result)
     * - Blocks: data-stat="blocks" from *_defense tables
     */
    private void globalKeywordSearch(MatchData match) {
        try {
            String script = "var result = { " +
                    "  bigChancesHome: '0', bigChancesAway: '0', " +
                    "  woodworkHome: '0', woodworkAway: '0', " +
                    "  throughBallsHome: '0', throughBallsAway: '0', " +
                    "  blocksHome: '0', blocksAway: '0', " +
                    "  touchesHome: '0', touchesAway: '0', " +
                    "  touchesOppBoxHome: '0', touchesOppBoxAway: '0', " +
                    "  dribblesHome: '0', dribblesAway: '0', " +
                    "  dribblesCompletedHome: '0', dribblesCompletedAway: '0' " +
                    "}; " +

                    // ========================================================
                    // PHYSICAL STATS: From possession tables
                    // data-stat: touches, touches_att_pen_area, dribbles, dribbles_completed
                    // ========================================================
                    "var possessionTables = document.querySelectorAll('table[id*=\"possession\"]'); " +
                    "var touchesVals = [], touchesOppVals = [], dribblesVals = [], dribblesCompVals = []; " +
                    "possessionTables.forEach(function(table) { " +
                    "  var tfoot = table.querySelector('tfoot'); " +
                    "  if (tfoot) { " +
                    "    var touchesCell = tfoot.querySelector('td[data-stat=\"touches\"]'); " +
                    "    var touchesOppCell = tfoot.querySelector('td[data-stat=\"touches_att_pen_area\"]'); " +
                    "    var dribblesCell = tfoot.querySelector('td[data-stat=\"dribbles\"]'); " +
                    "    var dribblesCompCell = tfoot.querySelector('td[data-stat=\"dribbles_completed\"]'); " +
                    "    if (touchesCell) touchesVals.push(touchesCell.textContent.trim() || '0'); " +
                    "    if (touchesOppCell) touchesOppVals.push(touchesOppCell.textContent.trim() || '0'); " +
                    "    if (dribblesCell) dribblesVals.push(dribblesCell.textContent.trim() || '0'); " +
                    "    if (dribblesCompCell) dribblesCompVals.push(dribblesCompCell.textContent.trim() || '0'); " +
                    "  } " +
                    "}); " +
                    "if (touchesVals.length >= 2) { result.touchesHome = touchesVals[0]; result.touchesAway = touchesVals[1]; } "
                    +
                    "if (touchesOppVals.length >= 2) { result.touchesOppBoxHome = touchesOppVals[0]; result.touchesOppBoxAway = touchesOppVals[1]; } "
                    +
                    "if (dribblesVals.length >= 2) { result.dribblesHome = dribblesVals[0]; result.dribblesAway = dribblesVals[1]; } "
                    +
                    "if (dribblesCompVals.length >= 2) { result.dribblesCompletedHome = dribblesCompVals[0]; result.dribblesCompletedAway = dribblesCompVals[1]; } "
                    +

                    // ========================================================
                    // THROUGH BALLS: From passing_types tables
                    // CORRECT: data-stat="through_balls" (NOT thru_balls)
                    // ========================================================
                    "var passingTables = document.querySelectorAll('table[id*=\"passing_types\"]'); " +
                    "var tbValues = []; " +
                    "passingTables.forEach(function(table) { " +
                    "  var tfoot = table.querySelector('tfoot'); " +
                    "  if (tfoot) { " +
                    "    var tbCell = tfoot.querySelector('td[data-stat=\"through_balls\"]'); " +
                    "    if (tbCell) { " +
                    "      tbValues.push(tbCell.textContent.trim() || '0'); " +
                    "    } " +
                    "  } " +
                    "}); " +
                    "if (tbValues.length >= 2) { " +
                    "  result.throughBallsHome = tbValues[0]; " +
                    "  result.throughBallsAway = tbValues[1]; " +
                    "} else if (tbValues.length === 1) { " +
                    "  result.throughBallsHome = tbValues[0]; " +
                    "} " +

                    // ========================================================
                    // BLOCKS: From defense tables
                    // data-stat="blocks"
                    // ========================================================
                    "var defenseTables = document.querySelectorAll('table[id*=\"defense\"]'); " +
                    "var blocksValues = []; " +
                    "defenseTables.forEach(function(table) { " +
                    "  var tfoot = table.querySelector('tfoot'); " +
                    "  if (tfoot) { " +
                    "    var blocksCell = tfoot.querySelector('td[data-stat=\"blocks\"]'); " +
                    "    if (blocksCell) { " +
                    "      blocksValues.push(blocksCell.textContent.trim() || '0'); " +
                    "    } " +
                    "  } " +
                    "}); " +
                    "if (blocksValues.length >= 2) { " +
                    "  result.blocksHome = blocksValues[0]; " +
                    "  result.blocksAway = blocksValues[1]; " +
                    "} else if (blocksValues.length === 1) { " +
                    "  result.blocksHome = blocksValues[0]; " +
                    "} " +

                    // ========================================================
                    // BIG CHANCES: Count shots with xG >= 0.35
                    // CORRECT: data-stat="xg_shot" (NOT xg)
                    // CORRECT: data-stat="team" (NOT squad)
                    // ========================================================
                    "var shotsTable = document.querySelector('#shots_all'); " +
                    "if (shotsTable) { " +
                    "  var homeTeamLink = document.querySelector('.scorebox > div:first-child strong a'); " +
                    "  var homeTeam = homeTeamLink ? homeTeamLink.textContent.toLowerCase() : ''; " +
                    "  var homeBigChances = 0; " +
                    "  var awayBigChances = 0; " +
                    "  var rows = shotsTable.querySelectorAll('tbody tr'); " +
                    "  rows.forEach(function(row) { " +
                    "    if (row.classList.contains('spacer') || row.classList.contains('thead')) return; " +
                    "    var xgCell = row.querySelector('td[data-stat=\"xg_shot\"]'); " +
                    "    var teamCell = row.querySelector('td[data-stat=\"team\"]'); " +
                    "    if (xgCell && teamCell) { " +
                    "      var xgVal = parseFloat(xgCell.textContent) || 0; " +
                    "      var team = teamCell.textContent.toLowerCase(); " +
                    "      if (xgVal >= 0.35) { " +
                    "        if (team.includes(homeTeam.substring(0,4)) || homeTeam.includes(team.substring(0,4))) { " +
                    "          homeBigChances++; " +
                    "        } else { " +
                    "          awayBigChances++; " +
                    "        } " +
                    "      } " +
                    "    } " +
                    "  }); " +
                    "  result.bigChancesHome = homeBigChances.toString(); " +
                    "  result.bigChancesAway = awayBigChances.toString(); " +
                    "} " +

                    // ========================================================
                    // WOODWORK: Check outcome column for "Post"
                    // CORRECT: data-stat="outcome" (NOT result)
                    // ========================================================
                    "if (shotsTable) { " +
                    "  var homeTeamLink = document.querySelector('.scorebox > div:first-child strong a'); " +
                    "  var homeTeam = homeTeamLink ? homeTeamLink.textContent.toLowerCase() : ''; " +
                    "  var homeWoodwork = 0; " +
                    "  var awayWoodwork = 0; " +
                    "  var rows = shotsTable.querySelectorAll('tbody tr'); " +
                    "  rows.forEach(function(row) { " +
                    "    if (row.classList.contains('spacer') || row.classList.contains('thead')) return; " +
                    "    var outcomeCell = row.querySelector('td[data-stat=\"outcome\"]'); " +
                    "    var teamCell = row.querySelector('td[data-stat=\"team\"]'); " +
                    "    if (outcomeCell && teamCell) { " +
                    "      var outcome = outcomeCell.textContent.toLowerCase(); " +
                    "      var team = teamCell.textContent.toLowerCase(); " +
                    "      if (outcome.includes('post') || outcome.includes('woodwork') || outcome.includes('bar')) { "
                    +
                    "        if (team.includes(homeTeam.substring(0,4)) || homeTeam.includes(team.substring(0,4))) { " +
                    "          homeWoodwork++; " +
                    "        } else { " +
                    "          awayWoodwork++; " +
                    "        } " +
                    "      } " +
                    "    } " +
                    "  }); " +
                    "  result.woodworkHome = homeWoodwork.toString(); " +
                    "  result.woodworkAway = awayWoodwork.toString(); " +
                    "} " +

                    "return result;";

            Object result = js.executeScript(script);
            if (result instanceof java.util.Map) {
                java.util.Map<?, ?> data = (java.util.Map<?, ?>) result;

                // Through Balls
                String foundThroughHome = getStr(data, "throughBallsHome");
                String foundThroughAway = getStr(data, "throughBallsAway");
                if ((match.homeThroughBalls.isEmpty() || match.homeThroughBalls.equals("0"))
                        && !foundThroughHome.equals("0") && !foundThroughHome.isEmpty()) {
                    match.homeThroughBalls = foundThroughHome;
                    match.awayThroughBalls = foundThroughAway;
                }

                // Blocks (NEW)
                String foundBlocksHome = getStr(data, "blocksHome");
                String foundBlocksAway = getStr(data, "blocksAway");
                if ((match.homeBlocks.isEmpty() || match.homeBlocks.equals("0"))
                        && !foundBlocksHome.equals("0") && !foundBlocksHome.isEmpty()) {
                    match.homeBlocks = foundBlocksHome;
                    match.awayBlocks = foundBlocksAway;
                }

                // Big Chances (xG >= 0.35)
                String foundBigChancesHome = getStr(data, "bigChancesHome");
                String foundBigChancesAway = getStr(data, "bigChancesAway");
                if ((match.homeBigChances.isEmpty() || match.homeBigChances.equals("0"))) {
                    match.homeBigChances = foundBigChancesHome.isEmpty() ? "0" : foundBigChancesHome;
                    match.awayBigChances = foundBigChancesAway.isEmpty() ? "0" : foundBigChancesAway;
                }

                // Woodwork (Post hits)
                String foundWoodworkHome = getStr(data, "woodworkHome");
                String foundWoodworkAway = getStr(data, "woodworkAway");
                if ((match.homeWoodwork.isEmpty() || match.homeWoodwork.equals("0"))) {
                    match.homeWoodwork = foundWoodworkHome.isEmpty() ? "0" : foundWoodworkHome;
                    match.awayWoodwork = foundWoodworkAway.isEmpty() ? "0" : foundWoodworkAway;
                }

                // === PHYSICAL STATS (NEW) ===
                // Touches
                String foundTouchesHome = getStr(data, "touchesHome");
                String foundTouchesAway = getStr(data, "touchesAway");
                if (match.homeTouches.isEmpty() && !foundTouchesHome.equals("0") && !foundTouchesHome.isEmpty()) {
                    match.homeTouches = foundTouchesHome;
                    match.awayTouches = foundTouchesAway;
                }

                // Touches in Opp Box
                String foundTouchesOppHome = getStr(data, "touchesOppBoxHome");
                String foundTouchesOppAway = getStr(data, "touchesOppBoxAway");
                if (match.homeTouchesOppBox.isEmpty() && !foundTouchesOppHome.equals("0")
                        && !foundTouchesOppHome.isEmpty()) {
                    match.homeTouchesOppBox = foundTouchesOppHome;
                    match.awayTouchesOppBox = foundTouchesOppAway;
                }

                // Dribbles
                String foundDribblesHome = getStr(data, "dribblesHome");
                String foundDribblesAway = getStr(data, "dribblesAway");
                if (match.homeDribbles.isEmpty() && !foundDribblesHome.equals("0") && !foundDribblesHome.isEmpty()) {
                    match.homeDribbles = foundDribblesHome;
                    match.awayDribbles = foundDribblesAway;
                }

                // Successful Dribbles
                String foundDribblesCompHome = getStr(data, "dribblesCompletedHome");
                String foundDribblesCompAway = getStr(data, "dribblesCompletedAway");
                if (match.homeDribblesCompleted.isEmpty() && !foundDribblesCompHome.equals("0")
                        && !foundDribblesCompHome.isEmpty()) {
                    match.homeDribblesCompleted = foundDribblesCompHome;
                    match.awayDribblesCompleted = foundDribblesCompAway;
                }
            }
        } catch (Exception e) {
            // Silently handle errors
        }
    }

    /**
     * Sleep helper
     */
    private void sleep(int ms) {
        try {
            Thread.sleep(ms);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }

    /**
     * Write all matches to Excel file with Dashboard Layout
     * - Sheet 1: Group Members
     * - Sheet 2-N: One sheet per match with dashboard layout
     */
    private void writeToExcel(List<MatchData> matches) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();

        // Create reusable styles
        CellStyle headerStyle = createHeaderStyle(wb);
        CellStyle titleStyle = createTitleStyle(wb);
        CellStyle scoreStyle = createScoreStyle(wb);
        CellStyle labelStyle = createLabelStyle(wb);
        CellStyle valueStyle = createValueStyle(wb);
        CellStyle centeredStyle = createCenteredStyle(wb);

        // ========== Sheet 1: Group Members ==========
        XSSFSheet coverSheet = wb.createSheet("Group Members");
        XSSFRow titleRow = coverSheet.createRow(0);
        XSSFCell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("EPL 2024-2025 Match Data - Group 33");
        titleCell.setCellStyle(titleStyle);
        coverSheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, 1));

        // Subtitle: UK Time Zone Notice
        XSSFRow subtitleRow = coverSheet.createRow(1);
        subtitleRow.createCell(0).setCellValue("All times displayed in UK Local Time (BST/GMT)");

        XSSFRow memberHeaderRow = coverSheet.createRow(3);
        memberHeaderRow.createCell(0).setCellValue("Student ID");
        memberHeaderRow.createCell(1).setCellValue("Name");
        memberHeaderRow.getCell(0).setCellStyle(headerStyle);
        memberHeaderRow.getCell(1).setCellStyle(headerStyle);

        // Student 1
        XSSFRow member1Row = coverSheet.createRow(4);
        member1Row.createCell(0).setCellValue("67011211012");
        member1Row.createCell(1).setCellValue("นายณภัทร ดีจันทึก");

        // Student 2
        XSSFRow member2Row = coverSheet.createRow(5);
        member2Row.createCell(0).setCellValue("67011211013");
        member2Row.createCell(1).setCellValue("นายณรงค์ฤทธิ์ พิมพ์แพทย์");

        // Student 3
        XSSFRow member3Row = coverSheet.createRow(6);
        member3Row.createCell(0).setCellValue("67011211014");
        member3Row.createCell(1).setCellValue("นางสาวณัฐชา จำเนียรสุข");

        // Student 4
        XSSFRow member4Row = coverSheet.createRow(7);
        member4Row.createCell(0).setCellValue("67011211017");
        member4Row.createCell(1).setCellValue("นายณัฐวุฒิ พละศักดิ์");

        coverSheet.autoSizeColumn(0);
        coverSheet.autoSizeColumn(1);

        // ========== Sheet 2-N: Match Sheets ==========
        int matchNum = 1;
        for (MatchData m : matches) {
            String sheetName = "Match_" + matchNum;
            XSSFSheet matchSheet = wb.createSheet(sheetName);

            createMatchDashboard(matchSheet, m, matchNum, headerStyle, titleStyle,
                    scoreStyle, labelStyle, valueStyle, centeredStyle);

            matchNum++;
        }

        // Write to file
        try (FileOutputStream out = new FileOutputStream(OUTPUT_FILE)) {
            wb.write(out);
        }
        wb.close();

        System.out.println("  Written " + matches.size() + " match sheets to " + OUTPUT_FILE);
    }

    /**
     * Create dashboard layout for a single match
     */
    private void createMatchDashboard(XSSFSheet sheet, MatchData m, int matchNum,
            CellStyle headerStyle, CellStyle titleStyle,
            CellStyle scoreStyle, CellStyle labelStyle,
            CellStyle valueStyle, CellStyle centeredStyle) {

        // ========== ZONE A: Overview (Cols A-C, Rows 1-15) ==========
        // Row 0: Match Title
        XSSFRow row0 = sheet.createRow(0);
        XSSFCell matchTitle = row0.createCell(0);
        matchTitle.setCellValue("Match #" + matchNum);
        matchTitle.setCellStyle(headerStyle);

        // Row 1: Teams vs Teams
        XSSFRow row1 = sheet.createRow(1);
        XSSFCell teamsCell = row1.createCell(0);
        teamsCell.setCellValue(m.homeTeam + "  vs  " + m.awayTeam);
        teamsCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(1, 1, 0, 3));

        // Row 2: "Score (FT)" Header (grey background)
        XSSFRow row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Score (FT)");
        row2.getCell(0).setCellStyle(headerStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(2, 2, 0, 1));

        // Row 3: FT Score Value (large bold centered)
        XSSFRow row3 = sheet.createRow(3);
        XSSFCell ftScoreCell = row3.createCell(0);
        ftScoreCell.setCellValue(m.fullTimeScore);
        ftScoreCell.setCellStyle(scoreStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(3, 3, 0, 1));

        // Row 4: "Score (HT)" Header (grey background)
        XSSFRow row4 = sheet.createRow(4);
        row4.createCell(0).setCellValue("Score (HT)");
        row4.getCell(0).setCellStyle(headerStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(4, 4, 0, 1));

        // Row 5: HT Score Value (bold centered)
        XSSFRow row5 = sheet.createRow(5);
        XSSFCell htScoreCell = row5.createCell(0);
        htScoreCell.setCellValue(m.halfTimeScore);
        htScoreCell.setCellStyle(scoreStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(5, 5, 0, 1));

        // Row 6: Goal Table Header (4 columns: A=Home Scorer, B=Min, C=Away Scorer,
        // D=Min)
        XSSFRow goalHeaderRow = sheet.createRow(6);
        goalHeaderRow.createCell(0).setCellValue("Home Scorer");
        goalHeaderRow.createCell(1).setCellValue("Min");
        goalHeaderRow.createCell(2).setCellValue("Away Scorer");
        goalHeaderRow.createCell(3).setCellValue("Min");
        goalHeaderRow.getCell(0).setCellStyle(headerStyle);
        goalHeaderRow.getCell(1).setCellStyle(headerStyle);
        goalHeaderRow.getCell(2).setCellStyle(headerStyle);
        goalHeaderRow.getCell(3).setCellStyle(headerStyle);

        // Row 7+: Goal Data - SEPARATE ROWS for Scorer and Assist (Professor's
        // Template)
        // Format: Row 1: Scorer Name | Min
        // Row 2: Assist Name | Min (Assist)
        String[] homeGoals = m.homeGoals.isEmpty() ? new String[0] : m.homeGoals.split(";");
        String[] awayGoals = m.awayGoals.isEmpty() ? new String[0] : m.awayGoals.split(";");

        // Build expanded lists with separate entries for scorers and assists
        java.util.List<String[]> homeEntries = new java.util.ArrayList<>();
        java.util.List<String[]> awayEntries = new java.util.ArrayList<>();

        // Process home goals
        for (String goal : homeGoals) {
            if (goal.isEmpty())
                continue;
            String[] parts = goal.split("\\|");
            String scorer = parts.length > 0 ? parts[0].trim() : "";
            String minute = parts.length > 1 ? parts[1].trim() : "";
            String assist = parts.length > 2 ? parts[2].trim() : "";

            // Add scorer row
            if (!scorer.isEmpty()) {
                homeEntries.add(new String[] { scorer, minute });
            }
            // Add assist row (separate line per Professor's Template)
            if (!assist.isEmpty()) {
                homeEntries.add(new String[] { assist, minute + " (Assist)" });
            }
        }

        // Process away goals
        for (String goal : awayGoals) {
            if (goal.isEmpty())
                continue;
            String[] parts = goal.split("\\|");
            String scorer = parts.length > 0 ? parts[0].trim() : "";
            String minute = parts.length > 1 ? parts[1].trim() : "";
            String assist = parts.length > 2 ? parts[2].trim() : "";

            // Add scorer row
            if (!scorer.isEmpty()) {
                awayEntries.add(new String[] { scorer, minute });
            }
            // Add assist row (separate line per Professor's Template)
            if (!assist.isEmpty()) {
                awayEntries.add(new String[] { assist, minute + " (Assist)" });
            }
        }

        int maxEntries = Math.max(homeEntries.size(), awayEntries.size());
        if (maxEntries == 0)
            maxEntries = 1;

        int currentGoalRow = 7;
        for (int i = 0; i < maxEntries; i++) {
            XSSFRow gRow = getOrCreateRow(sheet, currentGoalRow);

            // Home goal entry - Col A: Name, Col B: Minute (or Minute + Assist)
            if (i < homeEntries.size()) {
                String[] entry = homeEntries.get(i);
                gRow.createCell(0).setCellValue(entry[0]);
                gRow.createCell(1).setCellValue(entry[1]);
            }

            // Away goal entry - Col C: Name, Col D: Minute (or Minute + Assist)
            if (i < awayEntries.size()) {
                String[] entry = awayEntries.get(i);
                gRow.createCell(2).setCellValue(entry[0]);
                gRow.createCell(3).setCellValue(entry[1]);
            }
            currentGoalRow++;
        }

        // Cards Table Header (row after goals, minimum row 9)
        int cardsHeaderRow = Math.max(currentGoalRow + 1, 9);
        XSSFRow cardsHeader = getOrCreateRow(sheet, cardsHeaderRow);
        cardsHeader.createCell(0).setCellValue("Home Cards");
        cardsHeader.createCell(1).setCellValue("Min");
        cardsHeader.createCell(2).setCellValue("Away Cards");
        cardsHeader.createCell(3).setCellValue("Min");
        cardsHeader.getCell(0).setCellStyle(headerStyle);
        cardsHeader.getCell(1).setCellStyle(headerStyle);
        cardsHeader.getCell(2).setCellStyle(headerStyle);
        cardsHeader.getCell(3).setCellStyle(headerStyle);

        // Cards Data
        String[] homeCardsList = m.homeCards.isEmpty() ? new String[0] : m.homeCards.split(",");
        String[] awayCardsList = m.awayCards.isEmpty() ? new String[0] : m.awayCards.split(",");
        int maxCards = Math.max(homeCardsList.length, awayCardsList.length);

        for (int i = 0; i < maxCards && i < 5; i++) {
            XSSFRow cardRow = getOrCreateRow(sheet, cardsHeaderRow + 1 + i);
            // Home card - name and minute in separate columns
            if (i < homeCardsList.length && !homeCardsList[i].trim().isEmpty()) {
                String cardInfo = homeCardsList[i].trim();
                // Parse "Name Min" format
                String[] cardParts = cardInfo.split(" (?=\\d)");
                String name = cardParts.length > 0 ? cardParts[0].trim() : cardInfo;
                String min = cardParts.length > 1 ? cardParts[1].trim() : "";
                cardRow.createCell(0).setCellValue(name);
                cardRow.createCell(1).setCellValue(min);
            }
            // Away card
            if (i < awayCardsList.length && !awayCardsList[i].trim().isEmpty()) {
                String cardInfo = awayCardsList[i].trim();
                String[] cardParts = cardInfo.split(" (?=\\d)");
                String name = cardParts.length > 0 ? cardParts[0].trim() : cardInfo;
                String min = cardParts.length > 1 ? cardParts[1].trim() : "";
                cardRow.createCell(2).setCellValue(name);
                cardRow.createCell(3).setCellValue(min);
            }
        }

        // ========== ZONE B: Match Info & Officials (Cols E-F) ==========
        // Row 1: Header
        if (row1.getCell(4) == null)
            row1.createCell(4);
        row1.getCell(4).setCellValue("Match Info & Officials");
        row1.getCell(4).setCellStyle(headerStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(1, 1, 4, 5));

        // Row 2: Date
        row2.createCell(4).setCellValue("Date:");
        row2.createCell(5).setCellValue(m.dateTime);
        row2.getCell(4).setCellStyle(labelStyle);

        // Row 3: Stadium
        row3.createCell(4).setCellValue("Stadium:");
        row3.createCell(5).setCellValue(m.stadium);
        row3.getCell(4).setCellStyle(labelStyle);

        // Row 4: Attendance
        XSSFRow row4b = getOrCreateRow(sheet, 4);
        row4b.createCell(4).setCellValue("Attendance:");
        row4b.createCell(5).setCellValue(m.attendance);
        row4b.getCell(4).setCellStyle(labelStyle);

        // Row 5: Referee
        XSSFRow row5b = getOrCreateRow(sheet, 5);
        row5b.createCell(4).setCellValue("Referee:");
        row5b.createCell(5).setCellValue(m.referee);
        row5b.getCell(4).setCellStyle(labelStyle);

        // Row 6: Assistant Refs
        XSSFRow row6 = getOrCreateRow(sheet, 6);
        row6.createCell(4).setCellValue("AR:");
        row6.createCell(5).setCellValue(m.assistantRefs);
        row6.getCell(4).setCellStyle(labelStyle);

        // Row 7: 4th Official
        XSSFRow row7 = getOrCreateRow(sheet, 7);
        row7.createCell(4).setCellValue("4th:");
        row7.createCell(5).setCellValue(m.fourthOfficial);
        row7.getCell(4).setCellStyle(labelStyle);

        // Row 8: VAR
        XSSFRow row8x = getOrCreateRow(sheet, 8);
        row8x.createCell(4).setCellValue("VAR:");
        row8x.createCell(5).setCellValue(m.varOfficial);
        row8x.getCell(4).setCellStyle(labelStyle);

        // ========== ZONE C: Lineups (Cols H-I, Rows 1-15) ==========
        // Row 1: Header
        row1.createCell(7).setCellValue("Home XI");
        row1.createCell(8).setCellValue("Away XI");
        row1.getCell(7).setCellStyle(headerStyle);
        row1.getCell(8).setCellStyle(headerStyle);

        // Split lineup strings and write vertically
        String[] homePlayers = m.homeStartingXI.split(",");
        String[] awayPlayers = m.awayStartingXI.split(",");
        int maxPlayers = Math.max(homePlayers.length, awayPlayers.length);

        for (int i = 0; i < maxPlayers; i++) {
            XSSFRow playerRow = getOrCreateRow(sheet, 2 + i);

            if (i < homePlayers.length && !homePlayers[i].trim().isEmpty()) {
                XSSFCell homeCell = playerRow.createCell(7);
                // Remove leading numbers like "1.", "23." etc.
                String cleanName = homePlayers[i].trim().replaceAll("^\\d+\\.\\s*", "");
                homeCell.setCellValue(cleanName);
            }

            if (i < awayPlayers.length && !awayPlayers[i].trim().isEmpty()) {
                XSSFCell awayCell = playerRow.createCell(8);
                // Remove leading numbers like "1.", "23." etc.
                String cleanName = awayPlayers[i].trim().replaceAll("^\\d+\\.\\s*", "");
                awayCell.setCellValue(cleanName);
            }
        }

        // ========== Managers (between Starting XI and Substitutes) ==========
        int mgrRow = 2 + maxPlayers + 1; // 1 row gap after Starting XI
        XSSFRow mgrHeaderRow = getOrCreateRow(sheet, mgrRow);
        mgrHeaderRow.createCell(7).setCellValue("Manager");
        mgrHeaderRow.createCell(8).setCellValue("Manager");
        mgrHeaderRow.getCell(7).setCellStyle(headerStyle);
        mgrHeaderRow.getCell(8).setCellStyle(headerStyle);

        XSSFRow mgrDataRow = getOrCreateRow(sheet, mgrRow + 1);
        mgrDataRow.createCell(7).setCellValue(m.homeManager);
        mgrDataRow.createCell(8).setCellValue(m.awayManager);

        // ========== ZONE C2: Substitutes (below Managers) ==========
        int subsStartRow = mgrRow + 3; // 1 row gap after Managers
        XSSFRow subsHeaderRow = getOrCreateRow(sheet, subsStartRow);
        subsHeaderRow.createCell(7).setCellValue("Substitutes");
        subsHeaderRow.createCell(8).setCellValue("Substitutes");
        subsHeaderRow.getCell(7).setCellStyle(headerStyle);
        subsHeaderRow.getCell(8).setCellStyle(headerStyle);

        // Split substitutes and write vertically
        String[] homeSubs = m.homeSubstitutes.split(",");
        String[] awaySubs = m.awaySubstitutes.split(",");
        int maxSubs = Math.max(homeSubs.length, awaySubs.length);

        for (int i = 0; i < maxSubs && i < 9; i++) { // Max 9 subs
            XSSFRow subRow = getOrCreateRow(sheet, subsStartRow + 1 + i);

            if (i < homeSubs.length && !homeSubs[i].trim().isEmpty()) {
                XSSFCell homeSubCell = subRow.createCell(7);
                homeSubCell.setCellValue(homeSubs[i].trim());
            }

            if (i < awaySubs.length && !awaySubs[i].trim().isEmpty()) {
                XSSFCell awaySubCell = subRow.createCell(8);
                awaySubCell.setCellValue(awaySubs[i].trim());
            }
        }

        // ========== ZONE D: Stats Comparison (Cols K-M) ==========
        // Grouped by category as per professor's template

        // Row 1: Header
        row1.createCell(10).setCellValue("Home");
        row1.createCell(11).setCellValue("Stat");
        row1.createCell(12).setCellValue("Away");
        row1.getCell(10).setCellStyle(headerStyle);
        row1.getCell(11).setCellStyle(headerStyle);
        row1.getCell(12).setCellStyle(headerStyle);

        // === Zone D: Stats with CORRECT ORDER ===
        // ORDER: Top Stats → Attack → Possession → Defence → Physical → Discipline
        int statRow = 2;

        // === TOP STATS === (Header)
        XSSFRow topHeader = getOrCreateRow(sheet, statRow++);
        topHeader.createCell(10).setCellValue("");
        topHeader.createCell(11).setCellValue("--- TOP STATS ---");
        topHeader.createCell(12).setCellValue("");
        topHeader.getCell(11).setCellStyle(headerStyle);

        // Possession
        XSSFRow s1 = getOrCreateRow(sheet, statRow++);
        s1.createCell(10).setCellValue(m.homePossession);
        s1.createCell(11).setCellValue("Possession");
        s1.createCell(12).setCellValue(m.awayPossession);
        s1.getCell(10).setCellStyle(centeredStyle);
        s1.getCell(11).setCellStyle(labelStyle);
        s1.getCell(12).setCellStyle(centeredStyle);

        // xG
        XSSFRow s2 = getOrCreateRow(sheet, statRow++);
        s2.createCell(10).setCellValue(m.homeXG);
        s2.createCell(11).setCellValue("xG");
        s2.createCell(12).setCellValue(m.awayXG);
        s2.getCell(10).setCellStyle(centeredStyle);
        s2.getCell(11).setCellStyle(labelStyle);
        s2.getCell(12).setCellStyle(centeredStyle);

        // Total Shots
        XSSFRow s3 = getOrCreateRow(sheet, statRow++);
        s3.createCell(10).setCellValue(m.homeTotalShots);
        s3.createCell(11).setCellValue("Total Shots");
        s3.createCell(12).setCellValue(m.awayTotalShots);
        s3.getCell(10).setCellStyle(centeredStyle);
        s3.getCell(11).setCellStyle(labelStyle);
        s3.getCell(12).setCellStyle(centeredStyle);

        // Shots on Target
        XSSFRow s4 = getOrCreateRow(sheet, statRow++);
        s4.createCell(10).setCellValue(m.homeShotsOnTarget);
        s4.createCell(11).setCellValue("Shots on Target");
        s4.createCell(12).setCellValue(m.awayShotsOnTarget);
        s4.getCell(10).setCellStyle(centeredStyle);
        s4.getCell(11).setCellStyle(labelStyle);
        s4.getCell(12).setCellStyle(centeredStyle);

        // === ATTACK === (Header)
        XSSFRow attackHeader = getOrCreateRow(sheet, statRow++);
        attackHeader.createCell(10).setCellValue("");
        attackHeader.createCell(11).setCellValue("--- ATTACK ---");
        attackHeader.createCell(12).setCellValue("");
        attackHeader.getCell(11).setCellStyle(headerStyle);

        // Big Chances
        XSSFRow s5 = getOrCreateRow(sheet, statRow++);
        s5.createCell(10).setCellValue(m.homeBigChances.isEmpty() ? "0" : m.homeBigChances);
        s5.createCell(11).setCellValue("Big Chances");
        s5.createCell(12).setCellValue(m.awayBigChances.isEmpty() ? "0" : m.awayBigChances);
        s5.getCell(10).setCellStyle(centeredStyle);
        s5.getCell(11).setCellStyle(labelStyle);
        s5.getCell(12).setCellStyle(centeredStyle);

        // Hit Woodwork
        XSSFRow s6 = getOrCreateRow(sheet, statRow++);
        s6.createCell(10).setCellValue(m.homeWoodwork.isEmpty() ? "0" : m.homeWoodwork);
        s6.createCell(11).setCellValue("Hit Woodwork");
        s6.createCell(12).setCellValue(m.awayWoodwork.isEmpty() ? "0" : m.awayWoodwork);
        s6.getCell(10).setCellStyle(centeredStyle);
        s6.getCell(11).setCellStyle(labelStyle);
        s6.getCell(12).setCellStyle(centeredStyle);

        // Corners
        XSSFRow s7 = getOrCreateRow(sheet, statRow++);
        s7.createCell(10).setCellValue(m.homeCorners);
        s7.createCell(11).setCellValue("Corners");
        s7.createCell(12).setCellValue(m.awayCorners);
        s7.getCell(10).setCellStyle(centeredStyle);
        s7.getCell(11).setCellStyle(labelStyle);
        s7.getCell(12).setCellStyle(centeredStyle);

        // Crosses
        XSSFRow s8 = getOrCreateRow(sheet, statRow++);
        s8.createCell(10).setCellValue(m.homeCrosses.isEmpty() ? "0" : m.homeCrosses);
        s8.createCell(11).setCellValue("Crosses");
        s8.createCell(12).setCellValue(m.awayCrosses.isEmpty() ? "0" : m.awayCrosses);
        s8.getCell(10).setCellStyle(centeredStyle);
        s8.getCell(11).setCellStyle(labelStyle);
        s8.getCell(12).setCellStyle(centeredStyle);

        // === POSSESSION === (Header)
        XSSFRow possHeader = getOrCreateRow(sheet, statRow++);
        possHeader.createCell(10).setCellValue("");
        possHeader.createCell(11).setCellValue("--- POSSESSION ---");
        possHeader.createCell(12).setCellValue("");
        possHeader.getCell(11).setCellStyle(headerStyle);

        // Passes
        XSSFRow s9 = getOrCreateRow(sheet, statRow++);
        s9.createCell(10).setCellValue(m.homePasses);
        s9.createCell(11).setCellValue("Passes");
        s9.createCell(12).setCellValue(m.awayPasses);
        s9.getCell(10).setCellStyle(centeredStyle);
        s9.getCell(11).setCellStyle(labelStyle);
        s9.getCell(12).setCellStyle(centeredStyle);

        // Long Balls
        XSSFRow s10 = getOrCreateRow(sheet, statRow++);
        s10.createCell(10).setCellValue(m.homeLongBalls.isEmpty() ? "0" : m.homeLongBalls);
        s10.createCell(11).setCellValue("Long Balls");
        s10.createCell(12).setCellValue(m.awayLongBalls.isEmpty() ? "0" : m.awayLongBalls);
        s10.getCell(10).setCellStyle(centeredStyle);
        s10.getCell(11).setCellStyle(labelStyle);
        s10.getCell(12).setCellStyle(centeredStyle);

        // Through Balls
        XSSFRow s11 = getOrCreateRow(sheet, statRow++);
        s11.createCell(10).setCellValue(m.homeThroughBalls.isEmpty() ? "0" : m.homeThroughBalls);
        s11.createCell(11).setCellValue("Through Balls");
        s11.createCell(12).setCellValue(m.awayThroughBalls.isEmpty() ? "0" : m.awayThroughBalls);
        s11.getCell(10).setCellStyle(centeredStyle);
        s11.getCell(11).setCellStyle(labelStyle);
        s11.getCell(12).setCellStyle(centeredStyle);

        // Touches
        XSSFRow s12 = getOrCreateRow(sheet, statRow++);
        s12.createCell(10).setCellValue(m.homeTouches.isEmpty() ? "0" : m.homeTouches);
        s12.createCell(11).setCellValue("Touches");
        s12.createCell(12).setCellValue(m.awayTouches.isEmpty() ? "0" : m.awayTouches);
        s12.getCell(10).setCellStyle(centeredStyle);
        s12.getCell(11).setCellStyle(labelStyle);
        s12.getCell(12).setCellStyle(centeredStyle);

        // Touches in Opp Box
        XSSFRow s13 = getOrCreateRow(sheet, statRow++);
        s13.createCell(10).setCellValue(m.homeTouchesOppBox.isEmpty() ? "0" : m.homeTouchesOppBox);
        s13.createCell(11).setCellValue("Touches in Opp Box");
        s13.createCell(12).setCellValue(m.awayTouchesOppBox.isEmpty() ? "0" : m.awayTouchesOppBox);
        s13.getCell(10).setCellStyle(centeredStyle);
        s13.getCell(11).setCellStyle(labelStyle);
        s13.getCell(12).setCellStyle(centeredStyle);

        // === DEFENCE === (Header)
        XSSFRow defHeader = getOrCreateRow(sheet, statRow++);
        defHeader.createCell(10).setCellValue("");
        defHeader.createCell(11).setCellValue("--- DEFENCE ---");
        defHeader.createCell(12).setCellValue("");
        defHeader.getCell(11).setCellStyle(headerStyle);

        // Tackles Won
        XSSFRow s14 = getOrCreateRow(sheet, statRow++);
        s14.createCell(10).setCellValue(m.homeTackles);
        s14.createCell(11).setCellValue("Tackles Won");
        s14.createCell(12).setCellValue(m.awayTackles);
        s14.getCell(10).setCellStyle(centeredStyle);
        s14.getCell(11).setCellStyle(labelStyle);
        s14.getCell(12).setCellStyle(centeredStyle);

        // Blocks
        XSSFRow s15 = getOrCreateRow(sheet, statRow++);
        s15.createCell(10).setCellValue(m.homeBlocks.isEmpty() ? "0" : m.homeBlocks);
        s15.createCell(11).setCellValue("Blocks");
        s15.createCell(12).setCellValue(m.awayBlocks.isEmpty() ? "0" : m.awayBlocks);
        s15.getCell(10).setCellStyle(centeredStyle);
        s15.getCell(11).setCellStyle(labelStyle);
        s15.getCell(12).setCellStyle(centeredStyle);

        // Interceptions
        XSSFRow s16 = getOrCreateRow(sheet, statRow++);
        s16.createCell(10).setCellValue(m.homeInterceptions.isEmpty() ? "0" : m.homeInterceptions);
        s16.createCell(11).setCellValue("Interceptions");
        s16.createCell(12).setCellValue(m.awayInterceptions.isEmpty() ? "0" : m.awayInterceptions);
        s16.getCell(10).setCellStyle(centeredStyle);
        s16.getCell(11).setCellStyle(labelStyle);
        s16.getCell(12).setCellStyle(centeredStyle);

        // Clearances
        XSSFRow s17 = getOrCreateRow(sheet, statRow++);
        s17.createCell(10).setCellValue(m.homeClearances.isEmpty() ? "0" : m.homeClearances);
        s17.createCell(11).setCellValue("Clearances");
        s17.createCell(12).setCellValue(m.awayClearances.isEmpty() ? "0" : m.awayClearances);
        s17.getCell(10).setCellStyle(centeredStyle);
        s17.getCell(11).setCellStyle(labelStyle);
        s17.getCell(12).setCellStyle(centeredStyle);

        // Saves
        XSSFRow s18 = getOrCreateRow(sheet, statRow++);
        s18.createCell(10).setCellValue(m.homeSaves);
        s18.createCell(11).setCellValue("Saves");
        s18.createCell(12).setCellValue(m.awaySaves);
        s18.getCell(10).setCellStyle(centeredStyle);
        s18.getCell(11).setCellStyle(labelStyle);
        s18.getCell(12).setCellStyle(centeredStyle);

        // === PHYSICAL === (Header) - NEW SECTION
        XSSFRow physHeader = getOrCreateRow(sheet, statRow++);
        physHeader.createCell(10).setCellValue("");
        physHeader.createCell(11).setCellValue("--- PHYSICAL ---");
        physHeader.createCell(12).setCellValue("");
        physHeader.getCell(11).setCellStyle(headerStyle);

        // Dribbles
        XSSFRow s19 = getOrCreateRow(sheet, statRow++);
        s19.createCell(10).setCellValue(m.homeDribbles.isEmpty() ? "0" : m.homeDribbles);
        s19.createCell(11).setCellValue("Dribbles");
        s19.createCell(12).setCellValue(m.awayDribbles.isEmpty() ? "0" : m.awayDribbles);
        s19.getCell(10).setCellStyle(centeredStyle);
        s19.getCell(11).setCellStyle(labelStyle);
        s19.getCell(12).setCellStyle(centeredStyle);

        // Successful Dribbles
        XSSFRow s20 = getOrCreateRow(sheet, statRow++);
        s20.createCell(10).setCellValue(m.homeDribblesCompleted.isEmpty() ? "0" : m.homeDribblesCompleted);
        s20.createCell(11).setCellValue("Successful Dribbles");
        s20.createCell(12).setCellValue(m.awayDribblesCompleted.isEmpty() ? "0" : m.awayDribblesCompleted);
        s20.getCell(10).setCellStyle(centeredStyle);
        s20.getCell(11).setCellStyle(labelStyle);
        s20.getCell(12).setCellStyle(centeredStyle);

        // Aerial Duels Won
        XSSFRow s21 = getOrCreateRow(sheet, statRow++);
        s21.createCell(10).setCellValue(m.homeAerials.isEmpty() ? "0" : m.homeAerials);
        s21.createCell(11).setCellValue("Aerial Duels Won");
        s21.createCell(12).setCellValue(m.awayAerials.isEmpty() ? "0" : m.awayAerials);
        s21.getCell(10).setCellStyle(centeredStyle);
        s21.getCell(11).setCellStyle(labelStyle);
        s21.getCell(12).setCellStyle(centeredStyle);

        // Distance Covered
        XSSFRow s22 = getOrCreateRow(sheet, statRow++);
        s22.createCell(10).setCellValue("N/A");
        s22.createCell(11).setCellValue("Distance Covered");
        s22.createCell(12).setCellValue("N/A");
        s22.getCell(10).setCellStyle(centeredStyle);
        s22.getCell(11).setCellStyle(labelStyle);
        s22.getCell(12).setCellStyle(centeredStyle);

        // === DISCIPLINE === (Header) - MOVED TO END
        XSSFRow discHeader = getOrCreateRow(sheet, statRow++);
        discHeader.createCell(10).setCellValue("");
        discHeader.createCell(11).setCellValue("--- DISCIPLINE ---");
        discHeader.createCell(12).setCellValue("");
        discHeader.getCell(11).setCellStyle(headerStyle);

        // Yellow Cards
        XSSFRow s23 = getOrCreateRow(sheet, statRow++);
        s23.createCell(10).setCellValue(m.homeYellowCards);
        s23.createCell(11).setCellValue("Yellow Cards");
        s23.createCell(12).setCellValue(m.awayYellowCards);
        s23.getCell(10).setCellStyle(centeredStyle);
        s23.getCell(11).setCellStyle(labelStyle);
        s23.getCell(12).setCellStyle(centeredStyle);

        // Red Cards
        XSSFRow s24 = getOrCreateRow(sheet, statRow++);
        s24.createCell(10).setCellValue(m.homeRedCards);
        s24.createCell(11).setCellValue("Red Cards");
        s24.createCell(12).setCellValue(m.awayRedCards);
        s24.getCell(10).setCellStyle(centeredStyle);
        s24.getCell(11).setCellStyle(labelStyle);
        s24.getCell(12).setCellStyle(centeredStyle);

        // Fouls
        XSSFRow s25 = getOrCreateRow(sheet, statRow++);
        s25.createCell(10).setCellValue(m.homeFouls);
        s25.createCell(11).setCellValue("Fouls");
        s25.createCell(12).setCellValue(m.awayFouls);
        s25.getCell(10).setCellStyle(centeredStyle);
        s25.getCell(11).setCellStyle(labelStyle);
        s25.getCell(12).setCellStyle(centeredStyle);

        // Offsides
        XSSFRow s26 = getOrCreateRow(sheet, statRow++);
        s26.createCell(10).setCellValue(m.homeOffsides);
        s26.createCell(11).setCellValue("Offsides");
        s26.createCell(12).setCellValue(m.awayOffsides);
        s26.getCell(10).setCellStyle(centeredStyle);
        s26.getCell(11).setCellStyle(labelStyle);
        s26.getCell(12).setCellStyle(centeredStyle);

        // Auto-size ALL columns (A to P = indices 0 to 15)
        for (int i = 0; i <= 15; i++) {
            sheet.autoSizeColumn(i);
        }

        // Set minimum widths for columns with long text
        if (sheet.getColumnWidth(0) < 8000)
            sheet.setColumnWidth(0, 8000); // Col A - Scorer names
        if (sheet.getColumnWidth(2) < 8000)
            sheet.setColumnWidth(2, 8000); // Col C - Away Scorer names
        sheet.setColumnWidth(5, 8000); // Col F (Match Info values)
        sheet.setColumnWidth(7, 6000); // Col H (Home XI)
        sheet.setColumnWidth(8, 6000); // Col I (Away XI)
        sheet.setColumnWidth(11, 5000); // Col L (Stat labels)
    }

    /**
     * Helper: Get or create row
     */
    private XSSFRow getOrCreateRow(XSSFSheet sheet, int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        return row;
    }

    /**
     * Create header style (dark blue background, white bold text)
     */
    private CellStyle createHeaderStyle(XSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Create title style (large bold text)
     */
    private CellStyle createTitleStyle(XSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        return style;
    }

    /**
     * Create score style (extra large bold centered)
     */
    private CellStyle createScoreStyle(XSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 24);
        font.setColor(IndexedColors.DARK_RED.getIndex());
        style.setFont(font);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Create label style (bold text)
     */
    private CellStyle createLabelStyle(XSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * Create value style (normal text)
     */
    private CellStyle createValueStyle(XSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        return style;
    }

    /**
     * Create centered style
     */
    private CellStyle createCenteredStyle(XSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * Data class to hold match information
     */
    private static class MatchData {
        String homeTeam = "";
        String awayTeam = "";
        String fullTimeScore = "";
        String halfTimeScore = ""; // HT Score e.g., "0-0"

        // Goal Events: format "Scorer|Min|Assist;" for each goal
        String homeGoals = ""; // e.g., "Zirkzee|87'|Garnacho"
        String awayGoals = "";

        // Cards split by team: format "Player Min, Player Min"
        String homeCards = ""; // e.g., "Mount 18', Maguire 40'"
        String awayCards = ""; // e.g., "Bassey 25', Pereira 70'"

        String dateTime = "";
        String stadium = "";
        String attendance = "";

        // Match Officials
        String referee = "";
        String assistantRefs = ""; // e.g., "Ian Hussin, Simon Bennett"
        String fourthOfficial = "";
        String varOfficial = "";

        // Managers
        String homeManager = "";
        String awayManager = "";

        // Starting XI
        String homeStartingXI = "";
        String awayStartingXI = "";

        // Substitutes with minutes - format: "Player Name (Min')"
        String homeSubstitutes = "";
        String awaySubstitutes = "";

        // Basic Stats
        String homePossession = "";
        String awayPossession = "";
        String homeTotalShots = "";
        String awayTotalShots = "";
        String homeShotsOnTarget = "";
        String awayShotsOnTarget = "";

        // Advanced Stats
        String homeXG = ""; // Expected Goals
        String awayXG = "";
        String homePasses = ""; // Completed of Total (e.g., "430 of 524")
        String awayPasses = "";
        String homeTackles = "";
        String awayTackles = "";

        // Discipline Stats
        String homeCorners = "";
        String awayCorners = "";
        String homeFouls = "";
        String awayFouls = "";
        String homeOffsides = "";
        String awayOffsides = "";
        String homeYellowCards = "";
        String awayYellowCards = "";
        String homeRedCards = "";
        String awayRedCards = "";

        // Defence Stats
        String homeSaves = "";
        String awaySaves = "";
        String homeClearances = "";
        String awayClearances = "";
        String homeInterceptions = "";
        String awayInterceptions = "";
        String homeBlocks = "";
        String awayBlocks = "";
        String homeAerials = ""; // Aerial Duels Won
        String awayAerials = "";

        // Attack Stats
        String homeBigChances = "";
        String awayBigChances = "";
        String homeWoodwork = ""; // Hit Woodwork
        String awayWoodwork = "";
        String homeCrosses = "";
        String awayCrosses = "";

        // Possession Stats
        String homeThroughBalls = "";
        String awayThroughBalls = "";
        String homeLongBalls = "";
        String awayLongBalls = "";

        // Physical Stats (NEW)
        String homeTouches = "";
        String awayTouches = "";
        String homeTouchesOppBox = ""; // Touches in Opp Penalty Area
        String awayTouchesOppBox = "";
        String homeDribbles = "";
        String awayDribbles = "";
        String homeDribblesCompleted = "";
        String awayDribblesCompleted = "";
    }
}
