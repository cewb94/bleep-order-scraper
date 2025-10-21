package com.dj.bleeporderscraper;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.*;

import org.apache.poi.ss.usermodel.Cell;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BleepOrderScraper {

    // 👉 Update these to your environment
    private static final String ORDERS_FILE = "C:\\Users\\Yaqoub Alshatti\\Projects\\DJ\\DJ Mapper\\imports\\MF_Orders_unscraped_aof_20-Oct-25.csv";
    private static final String OUTPUT_CSV  = "C:\\Users\\Yaqoub Alshatti\\Projects\\DJ\\DJ Mapper\\exports\\Bleep Orders\\bleep_item_export.csv";

    private static final String BLEEP_EMAIL    = "abdullahalawadhi1994@gmail.com";
    private static final String BLEEP_PASSWORD = "";

    private static final String LOGIN_URL = "https://dashboard.trybleep.com/auth/login";
    // NOTE: Base orders URL without any prefix; we append the exact H value
    private static final String ORDER_URL = "https://dashboard.trybleep.com/dashboard/orders/";

    private static final Duration PAGE_WAIT  = Duration.ofSeconds(15);
    private static final Duration SMALL_WAIT = Duration.ofSeconds(5);

    // Keep the browser open after the run (set false to auto-close)
    private static final boolean KEEP_BROWSER_OPEN = true;

    public static void main(String[] args) throws Exception {
        // Read Customer Reference from column H (MyFatoorah export)
        List<String> orderIds = readCustomerRefs(ORDERS_FILE);
        if (orderIds.isEmpty()) {
            System.err.println("No Customer References found in column H. Exiting.");
            return;
        }
        System.out.println("Loaded " + orderIds.size() + " customer references from " + ORDERS_FILE);
        for (String id : orderIds) {
            System.out.println("  - " + id);
        }

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized"); // headful
        options.setExperimentalOption("excludeSwitches", Arrays.asList("enable-automation"));
        options.setExperimentalOption("useAutomationExtension", false);

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, PAGE_WAIT);

        List<String[]> outRows = new ArrayList<>();
        outRows.add(new String[]{"order_id", "order_number", "qty", "item_name"}); // CSV header

        try {
            login(driver, wait);             // includes a fixed 20s wait post-submit
            selectDesertJerky(driver, wait); // clicks the vendor card + waits another 20s

            for (String ordIdRaw : orderIds) {
                String ordId = ordIdRaw == null ? "" : ordIdRaw.trim();
                if (ordId.isEmpty()) continue;

                String url = ORDER_URL + ordId;
                System.out.println("\nProcessing id from column H: " + ordId);
                System.out.println("-> " + url);
                driver.get(url);

                // Wait for an order landmark to ensure SPA is done
                waitForOrderPage(wait);

                // Order number, e.g., "Order #EYSUD"
                String orderNumber = findOrderNumber(driver);
                if (orderNumber.isEmpty()) {
                    System.out.println("   [!] Could not find order number for " + ordId);
                } else {
                    System.out.println("   [i] " + orderNumber);
                }

                // Prefer left panel items
                List<Item> items = scrapeLeftItems(driver);
                if (items.isEmpty()) {
                    items = scrapeInvoiceItemsTable(driver);
                }

                if (items.isEmpty()) {
                    System.out.println("   [!] No items found for " + ordId);
                } else {
                    for (Item it : items) {
                        outRows.add(new String[]{ordId, orderNumber, String.valueOf(it.qty), it.name});
                        System.out.println("   [OK] " + it.qty + " x " + it.name);
                    }
                }
            }
        } finally {
            if (!KEEP_BROWSER_OPEN) {
                try { driver.quit(); } catch (Exception ignored) {}
            } else {
                System.out.println("\nBrowser left open (KEEP_BROWSER_OPEN = true).");
                System.out.println("Press ENTER in this console to close the browser and exit...");
                try { new java.util.Scanner(System.in).nextLine(); } catch (Exception ignored) {}
                try { driver.quit(); } catch (Exception ignored) {}
            }
        }

        writeCsv(OUTPUT_CSV, outRows);
        System.out.println("\nDone. Wrote: " + OUTPUT_CSV);
    }

    // ---------- Login with your exact Angular inputs ----------
    private static void login(WebDriver driver, WebDriverWait wait) {
        driver.get(LOGIN_URL);

        // Given HTML matches these selectors:
        By emailBox = By.cssSelector("input[placeholder='Email Address'][formcontrolname='email']");
        By passBox  = By.cssSelector("input[placeholder='Password'][type='password'][formcontrolname='password']");
        By submit   = By.cssSelector("button.btn-login");

        WebElement emailEl = wait.until(ExpectedConditions.elementToBeClickable(emailBox));
        WebElement passEl  = wait.until(ExpectedConditions.visibilityOfElementLocated(passBox));

        clearAndType(emailEl, BLEEP_EMAIL);
        clearAndType(passEl, BLEEP_PASSWORD);

        // Wait for button to become enabled (disabled attribute removed)
        wait.until((ExpectedCondition<Boolean>) d -> {
            try {
                WebElement btn = d.findElement(submit);
                String disabled = btn.getAttribute("disabled");
                return disabled == null || disabled.isEmpty();
            } catch (org.openqa.selenium.NoSuchElementException e) {
                return false;
            }
        });

        safeClick(driver, submit);

        // Hard wait 20s after submitting credentials (as requested)
        System.out.println("Waiting 20 seconds after login submit (headful)...");
        sleep(20_000);

        // Then wait for a post-login signal (best-effort)
        try {
            wait.until(ExpectedConditions.or(
                ExpectedConditions.urlContains("/dashboard"),
                ExpectedConditions.presenceOfElementLocated(By.cssSelector("app-sidebar, .menu-list.nav-links, [href='/dashboard/orders']"))
            ));
        } catch (Exception e) {
            // proceed – vendor list might still be showing
            System.out.println(String.valueOf(e));
        }
    }

    // ---------- Select the "Desert Jerky" company card then wait ----------
    private static void selectDesertJerky(WebDriver driver, WebDriverWait wait) {
        System.out.println("Looking for 'Desert Jerky' vendor card...");

        By cardByText = By.xpath(
            "//div[contains(@class,'shop-comp')]//div[contains(@class,'fs-18') and (normalize-space()='Desert Jerky' or normalize-space()='desertjerkykw.com')]/ancestor::div[contains(@class,'shop-comp')]"
        );

        try {
            WebElement card = wait.until(ExpectedConditions.presenceOfElementLocated(cardByText));
            WebElement clickable = null;
            try {
                clickable = card.findElement(By.cssSelector(".d-flex.justify-content-between.v-center"));
            } catch (org.openqa.selenium.NoSuchElementException ignored) {}

            if (clickable == null) clickable = card;

            try {
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", clickable);
            } catch (Exception ignored) {}
            sleep(150);

            safeClickElement(driver, clickable);

            // Wait the same amount of time after choosing the shop
            System.out.println("Vendor selected. Waiting 20 seconds for workspace to load...");
            sleep(20_000);

            // Optional: wait until dashboard chrome appears
            try {
                wait.until(ExpectedConditions.or(
                    ExpectedConditions.urlContains("/dashboard"),
                    ExpectedConditions.presenceOfElementLocated(By.cssSelector("app-sidebar, .menu-list.nav-links, [href='/dashboard/orders']"))
                ));
            } catch (Exception e) {
                System.out.println(String.valueOf(e));
            }
        } catch (Exception e) {
            System.out.println(String.valueOf(e));
        }
    }

    // ---------- After navigating to an order page ----------
    private static void waitForOrderPage(WebDriverWait wait) {
        wait.until(ExpectedConditions.or(
            // generalized to any receipt id prefix
            ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.invoice-container[id^='receipt_']")),
            ExpectedConditions.presenceOfElementLocated(By.cssSelector(".scrollable-items")),
            ExpectedConditions.presenceOfElementLocated(By.cssSelector(".items-table tbody tr"))
        ));
        // Small settle time for Angular change detection
        sleep(150);
    }

    private static String findOrderNumber(WebDriver driver) {
        // <span>Order #EYSUD</span>
        List<By> locators = Arrays.asList(
            By.xpath("//span[starts-with(normalize-space(.), 'Order #')]"),
            By.cssSelector(".order-id span"),
            By.xpath("//div[contains(@class,'order-header')]//span[starts-with(normalize-space(.), 'Order #')]")
        );
        for (By by : locators) {
            try {
                WebElement el = driver.findElement(by);
                String txt = el.getText().trim();
                if (txt.startsWith("Order #")) return txt;
            } catch (org.openqa.selenium.NoSuchElementException ignored) {}
        }
        return "";
    }

    // ---------- Primary: left “scrollable-items” ----------
    private static List<Item> scrapeLeftItems(WebDriver driver) {
        // <div class="item-qty"><span>x5</span></div> + <div class="item-name">Teriyaki</div>
        List<Item> out = new ArrayList<>();
        List<WebElement> blocks = driver.findElements(By.cssSelector(".scrollable-items .order-item .item-content"));
        for (WebElement b : blocks) {
            String qtyText  = textOrEmpty(b, By.cssSelector(".item-qty span")).trim(); // "x5"
            String nameText = textOrEmpty(b, By.cssSelector(".item-name")).trim();     // "Teriyaki"
            if (nameText.isEmpty()) continue;
            out.add(new Item(nameText, parseQty(qtyText)));
        }
        return out;
    }

    // ---------- Fallback: invoice table ----------
    private static List<Item> scrapeInvoiceItemsTable(WebDriver driver) {
        List<Item> out = new ArrayList<>();
        List<WebElement> rows = driver.findElements(By.cssSelector(".items-table tbody tr.item-row"));
        for (WebElement tr : rows) {
            String name = textOrEmpty(tr, By.cssSelector(".item-details .item-name")).trim();
            String qtyS = textOrEmpty(tr, By.cssSelector(".item-quantity")).trim();
            if (name.isEmpty()) continue;
            out.add(new Item(name, safeParseInt(qtyS, 0)));
        }
        return out;
    }

    // ---------- Helpers ----------
    private static void clearAndType(WebElement el, String text) {
        try { el.click(); } catch (Exception ignored) {}
        el.sendKeys(Keys.chord(Keys.CONTROL, "a"));
        el.sendKeys(Keys.DELETE);
        el.sendKeys(text);
    }

    private static void safeClick(WebDriver driver, By by) {
        WebDriverWait shortWait = new WebDriverWait(driver, SMALL_WAIT);
        try {
            WebElement el = shortWait.until(ExpectedConditions.elementToBeClickable(by));
            try {
                el.click();
            } catch (Exception clickIntercept) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", el);
            }
        } catch (Exception e) {
            try {
                WebElement el = driver.findElement(by);
                new Actions(driver).moveToElement(el).click().perform();
            } catch (Exception ignore) {}
        }
    }

    private static void safeClickElement(WebDriver driver, WebElement el) {
        try {
            new WebDriverWait(driver, SMALL_WAIT).until(ExpectedConditions.elementToBeClickable(el));
            try {
                el.click();
            } catch (Exception clickIntercept) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", el);
            }
        } catch (Exception e) {
            try {
                new Actions(driver).moveToElement(el).click().perform();
            } catch (Exception ignore) {}
        }
    }

    private static String textOrEmpty(WebElement root, By by) {
        try {
            return root.findElement(by).getText();
        } catch (org.openqa.selenium.NoSuchElementException e) {
            return "";
        }
    }

    private static int parseQty(String qty) {
        if (qty == null) return 0;
        String cleaned = qty.trim().replaceFirst("^[xX]\\s*", ""); // remove leading x
        return safeParseInt(cleaned, 0);
        }

    private static int safeParseInt(String s, int dflt) {
        try {
            return Integer.parseInt(s.replaceAll("[^0-9-]", ""));
        } catch (Exception e) {
            return dflt;
        }
    }

    // ---------- Column H reader ----------
    // Reads Customer Reference values from column H (index 7) of a MyFatoorah export.
    // Supports .xlsx (Apache POI) and .csv. Also tries a header named "Customer Reference" if present.
    private static List<String> readCustomerRefs(String path) throws IOException {
        String lower = path.toLowerCase(Locale.ROOT);
        if (lower.endsWith(".xlsx")) {
            return readCustomerRefsFromXlsx(path);
        } else if (lower.endsWith(".csv")) {
            return readCustomerRefsFromCsv(path);
        } else {
            throw new IOException("Unsupported file type for " + path + " (expected .xlsx or .csv)");
        }
    }

    private static List<String> readCustomerRefsFromXlsx(String path) throws IOException {
        final int H_COL = 7; // 0-based => column H
        List<String> refs = new ArrayList<>();
        try (InputStream in = new FileInputStream(path);
             Workbook wb = new XSSFWorkbook(in)) {

            Sheet sheet = wb.getSheetAt(0);
            if (sheet == null) return Collections.emptyList();

            Iterator<Row> it = sheet.rowIterator();
            if (!it.hasNext()) return Collections.emptyList();

            Row first = it.next();
            // Check if first row is a header with "Customer Reference"
            Map<String, Integer> headerIndex = new HashMap<>();
            for (Cell cell : first) {
                headerIndex.put(getCellString(cell).trim().toLowerCase(Locale.ROOT), cell.getColumnIndex());
            }
            Integer headerCol = null;
            for (String key : headerIndex.keySet()) {
                if (key.equals("customer reference") || key.equals("customer_reference")) {
                    headerCol = headerIndex.get(key);
                    break;
                }
            }

            // If first row was header, process remaining rows; otherwise include first row as data using fixed H column
            if (headerCol != null) {
                while (it.hasNext()) {
                    Row r = it.next();
                    Cell c = r.getCell(headerCol, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String v = getCellString(c).trim();
                    addIfValidRef(refs, v);
                }
            } else {
                // Include the first row's H value as data
                Cell cFirst = first.getCell(H_COL, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                addIfValidRef(refs, getCellString(cFirst).trim());
                while (it.hasNext()) {
                    Row r = it.next();
                    Cell c = r.getCell(H_COL, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    addIfValidRef(refs, getCellString(c).trim());
                }
            }
        }
        // de-duplicate while preserving order
        return refs.stream().filter(s -> !s.isEmpty()).distinct().collect(Collectors.toList());
    }

    private static List<String> readCustomerRefsFromCsv(String path) throws IOException {
        final int H_COL = 7; // 0-based => column H
        List<String> refs = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(path), StandardCharsets.UTF_8))) {
            String line;
            boolean headerChecked = false;
            Integer customerRefCol = null; // if header names present
            while ((line = br.readLine()) != null) {
                // Simple CSV split (handles quotes minimally)
                List<String> cols = parseCsvLine(line);
                if (!headerChecked) {
                    headerChecked = true;
                    // Try header detection
                    for (int i = 0; i < cols.size(); i++) {
                        String h = cols.get(i).trim().toLowerCase(Locale.ROOT);
                        if (h.equals("customer reference") || h.equals("customer_reference")) {
                            customerRefCol = i;
                            break;
                        }
                    }
                    if (customerRefCol != null) continue; // header line consumed, move to next
                }

                int idx = (customerRefCol != null) ? customerRefCol : H_COL;
                if (idx < cols.size()) {
                    String v = cols.get(idx).trim();
                    addIfValidRef(refs, v);
                }
            }
        }
        return refs.stream().filter(s -> !s.isEmpty()).distinct().collect(Collectors.toList());
    }

    private static void addIfValidRef(List<String> out, String v) {
        if (v == null) return;
        v = v.trim();
        if (v.isEmpty()) return;
        // Do NOT strip/normalize; use exactly what’s in column H
        out.add(v);
    }

    // Minimal CSV parser for common cases (quotes + commas)
    private static List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        if (line == null) return out;
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char ch = line.charAt(i);
            if (inQuotes) {
                if (ch == '\"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '\"') {
                        cur.append('\"'); // escaped quote
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cur.append(ch);
                }
            } else {
                if (ch == '\"') {
                    inQuotes = true;
                } else if (ch == ',') {
                    out.add(cur.toString());
                    cur.setLength(0);
                } else {
                    cur.append(ch);
                }
            }
        }
        out.add(cur.toString());
        return out;
    }

    private static String normalize(Cell c) {
        return getCellString(c).trim().toLowerCase();
    }

    private static String getCellString(Cell c) {
        if (c == null) return "";
        switch (c.getCellType()) {
            case STRING:  return c.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c)) return c.getDateCellValue().toString();
                double d = c.getNumericCellValue();
                long l = (long) d;
                return (Math.abs(d - l) < 1e-9) ? String.valueOf(l) : String.valueOf(d);
            case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
            case FORMULA:
                try { return c.getStringCellValue(); } catch (Exception e) { return String.valueOf(c.getNumericCellValue()); }
            default:      return "";
        }
    }

    private static void writeCsv(String path, List<String[]> rows) throws IOException {
        try (Writer w = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(path), StandardCharsets.UTF_8))) {
            for (String[] row : rows) {
                w.write(csvLine(row));
                w.write("\n");
            }
        }
    }

    private static String csvLine(String[] fields) {
        return Arrays.stream(fields)
            .map(BleepOrderScraper::csvEscape)
            .collect(Collectors.joining(","));
    }

    private static String csvEscape(String s) {
        if (s == null) s = "";
        boolean needQuotes = s.contains(",") || s.contains("\"") || s.contains("\n") || s.contains("\r");
        String v = s.replace("\"", "\"\"");
        return needQuotes ? "\"" + v + "\"" : v;
    }

    private static void sleep(long ms) {
        try { Thread.sleep(ms); } catch (InterruptedException ignored) {}
    }

    private static class Item {
        final String name;
        final int qty;
        Item(String name, int qty) { this.name = name; this.qty = qty; }
    }
}
