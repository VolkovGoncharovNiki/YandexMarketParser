using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

const string SearchResultsXPath = "//*[@data-auto='SerpStatic']";
const string ProductCardXPath = "//article[@data-auto='searchOrganic']";
const string PopupButtonXPath =
    "//button[" +
    "normalize-space()='\u041f\u043e\u043d\u044f\u0442\u043d\u043e' or " +
    "normalize-space()='\u0425\u043e\u0440\u043e\u0448\u043e' or " +
    "normalize-space()='\u0417\u0430\u043a\u0440\u044b\u0442\u044c' or " +
    "normalize-space()='\u041d\u0435 \u0441\u0435\u0439\u0447\u0430\u0441' or " +
    "normalize-space()='\u041f\u0440\u0438\u043d\u044f\u0442\u044c' or " +
    "normalize-space()='\u0421\u043e\u0433\u043b\u0430\u0441\u0435\u043d' or " +
    "@id='gdpr-popup-v3-button-all' or " +
    "@id='gdpr-popup-v3-button-mandatory' or " +
    "normalize-space()='Allow all' or " +
    "normalize-space()='Allow essential cookies']";

const string MarketArticleLabel = "\u0410\u0440\u0442\u0438\u043a\u0443\u043b \u041c\u0430\u0440\u043a\u0435\u0442\u0430";
const string NoPriceText = "\u0426\u0435\u043d\u0430 \u043d\u0435 \u0443\u043a\u0430\u0437\u0430\u043d\u0430";
const string QueryPromptText = "\u0412\u0432\u0435\u0434\u0438\u0442\u0435 \u043f\u043e\u0438\u0441\u043a\u043e\u0432\u044b\u0439 \u0437\u0430\u043f\u0440\u043e\u0441: ";
const string QueryMissingText = "\u041f\u043e\u0438\u0441\u043a\u043e\u0432\u044b\u0439 \u0437\u0430\u043f\u0440\u043e\u0441 \u043d\u0435 \u0437\u0430\u0434\u0430\u043d.";
const string SearchReturnedNoProductsText = "\u041f\u043e\u0438\u0441\u043a \u043d\u0435 \u0432\u0435\u0440\u043d\u0443\u043b \u0442\u043e\u0432\u0430\u0440\u044b \u0434\u043b\u044f \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438.";
const string ConfigureFiltersPromptText = "\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u0442\u0435 \u0444\u0438\u043b\u044c\u0442\u0440\u044b \u0432 \u043e\u0442\u043a\u0440\u044b\u0442\u043e\u043c \u043e\u043a\u043d\u0435 \u0431\u0440\u0430\u0443\u0437\u0435\u0440\u0430, \u0437\u0430\u0442\u0435\u043c \u043d\u0430\u0436\u043c\u0438\u0442\u0435 Enter \u0432 \u043a\u043e\u043d\u0441\u043e\u043b\u0438 \u0434\u043b\u044f \u043d\u0430\u0447\u0430\u043b\u0430 \u0441\u0431\u043e\u0440\u0430 \u0434\u0430\u043d\u043d\u044b\u0445...";
const string LoadingDetailsText = "\u041d\u0430\u0447\u0438\u043d\u0430\u044e \u0441\u0431\u043e\u0440 \u0434\u0435\u0442\u0430\u043b\u044c\u043d\u043e\u0439 \u0438\u043d\u0444\u043e\u0440\u043c\u0430\u0446\u0438\u0438 \u043f\u043e \u043a\u0430\u0440\u0442\u043e\u0447\u043a\u0430\u043c...";
const string CouldNotProcessText = "\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u0430\u0442\u044c \u0442\u043e\u0432\u0430\u0440";
const string ProcessedProductsText = "\u041e\u0431\u0440\u0430\u0431\u043e\u0442\u0430\u043d\u043e \u0442\u043e\u0432\u0430\u0440\u043e\u0432";
const string ExcelSavedText = "Excel-\u0444\u0430\u0439\u043b \u0441\u043e\u0445\u0440\u0430\u043d\u0435\u043d";
const string WaitTimeoutText = "\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0434\u043e\u0436\u0434\u0430\u0442\u044c\u0441\u044f \u0437\u0430\u0433\u0440\u0443\u0437\u043a\u0438 \u0441\u0442\u0440\u0430\u043d\u0438\u0446\u044b \u0438\u043b\u0438 \u043a\u0430\u0440\u0442\u043e\u0447\u0435\u043a \u0442\u043e\u0432\u0430\u0440\u043e\u0432.";
const string GenericErrorPrefix = "\u041f\u0440\u043e\u0438\u0437\u043e\u0448\u043b\u0430 \u043e\u0448\u0438\u0431\u043a\u0430";
const string SellerRatingBlockSelector = "[data-auto='rating-info-block'], [data-auto='widget-rating-stars-business']";
const string SellerLegalPopupSelector = "[id='/content/popup'], [data-apiary-widget-name='@light/Popup']";
const string SellerLegalAccordionSelector = "[data-auto='shop-info-juridical-accordion']";
const string SellerLegalInfoHeading = "\u042e\u0440\u0438\u0434\u0438\u0447\u0435\u0441\u043a\u0430\u044f \u0438\u043d\u0444\u043e\u0440\u043c\u0430\u0446\u0438\u044f";
const string InnLabel = "\u0418\u041d\u041d";
const string OgrnLabel = "\u041e\u0413\u0420\u041d";
const string OgrnipLabel = "\u041e\u0413\u0420\u041d\u0418\u041f";

const string ExtractProductsJs = """
    const cards = document.querySelectorAll("article[data-auto='searchOrganic']");
    const results = [];
    for (const card of cards) {
        const titleEl = card.querySelector("[data-auto='snippet-title']");
        if (!titleEl) continue;

        const title = (titleEl.innerText || titleEl.textContent || '').trim();
        if (!title) continue;

        let bestLink = null;
        const links = card.querySelectorAll("a[data-auto='snippet-link']");
        for (const link of links) {
            if (link.querySelector("[data-auto='snippet-title']")) {
                bestLink = link;
                break;
            }
        }

        if (!bestLink && links.length > 0) bestLink = links[0];
        if (!bestLink) continue;

        const href = (bestLink.href || '').trim();
        if (!href) continue;

        let price = '\u0426\u0435\u043d\u0430 \u043d\u0435 \u0443\u043a\u0430\u0437\u0430\u043d\u0430';
        const priceEl = card.querySelector("[data-auto='snippet-price-current']");
        if (priceEl) {
            const raw = (priceEl.innerText || priceEl.textContent || '').trim();
            if (raw) price = raw;
        }

        results.push({ title, url: href, price });
    }

    return results;
    """;

const string CountCardsJs =
    "return document.querySelectorAll(\"article[data-auto='searchOrganic']\").length;";

const string ScrollJs = """
    const placeholder = document.querySelector("[data-auto='infinityPlaceholder']");
    if (placeholder) {
        placeholder.scrollIntoView({ block: 'center', behavior: 'smooth' });
        return true;
    }

    window.scrollTo(0, document.body.scrollHeight);
    return false;
    """;

const string ExtractProductDetailsJs = """
    const normalize = value => (value || '')
        .replace(/\u00A0/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    const businessUrlPattern = /\/business--[^\/?#"'\\s]+\/\d+(?:\?[^"'\\s<]*)?/i;

    const result = {
        description: '',
        sellerName: '',
        sellerUrl: '',
        specs: []
    };

    const descriptionEl = document.querySelector("[data-auto='product-description']");
    if (descriptionEl) {
        result.description = normalize(descriptionEl.innerText || descriptionEl.textContent || '');
    }

    const specsRoot = document.querySelector("[data-auto='specs-list-fullExtended']");
    if (specsRoot) {
        const rows = specsRoot.querySelectorAll("[data-auto='product-spec']");
        for (const label of rows) {
            const name = normalize(label.innerText || label.textContent || '');
            if (!name) continue;

            let row = label.parentElement;
            while (row && row !== specsRoot) {
                if (row.querySelectorAll("[data-auto='product-spec']").length === 1 &&
                    row.querySelector(".eXP5k")) {
                    break;
                }
                row = row.parentElement;
            }

            if (!row) continue;

            const valueEl = row.querySelector(".eXP5k");
            const value = normalize(valueEl ? (valueEl.innerText || valueEl.textContent || '') : '');
            if (!value) continue;

            result.specs.push({ name, value });
        }
    }

    const businessLink = Array.from(document.querySelectorAll("a[href]"))
        .map(link => (link.href || '').trim())
        .find(href => businessUrlPattern.test(href));

    if (businessLink) {
        result.sellerUrl = businessLink;
    }

    const shopBlock = document.querySelector("[data-auto='shop-info-block']");
    if (shopBlock) {
        const link = shopBlock.querySelector("a[href]");
        if (link) {
            const title = shopBlock.querySelector("[data-auto='shop-info-title']");
            result.sellerName = normalize(title ? (title.innerText || title.textContent || '') : (link.innerText || link.textContent || ''));
            if (!result.sellerUrl) {
                result.sellerUrl = (link.href || '').trim();
            }
        }
    }

    return result;
    """;

Console.OutputEncoding = Encoding.UTF8;
Console.InputEncoding = Encoding.UTF8;

var query = ReadSearchQuery();
if (string.IsNullOrWhiteSpace(query))
{
    Console.WriteLine(QueryMissingText);
    return;
}

var searchUrl = $"https://market.yandex.ru/search?text={Uri.EscapeDataString(query)}";

ChromeOptions options = new();
options.AddArgument("--start-maximized");

using var driver = new ChromeDriver(options);
var js = (IJavaScriptExecutor)driver;
WebDriverWait wait = new(driver, TimeSpan.FromSeconds(30));

try
{
    driver.Navigate().GoToUrl(searchUrl);

    wait.Until(IsSearchResultsLoaded);

    Console.WriteLine();
    Console.WriteLine(ConfigureFiltersPromptText);
    Console.ReadLine();

    wait.Until(IsSearchResultsLoaded);

    ScrollUntilAllProductsLoaded(driver, js);

    var searchProducts = DeduplicateProducts(ExtractAllProductsViaJs(js));
    foreach (var product in searchProducts)
    {
        product.Query = query;
    }

    if (searchProducts.Count == 0)
    {
        Console.WriteLine(SearchReturnedNoProductsText);
        return;
    }

    Console.WriteLine();
    Console.WriteLine($"{LoadingDetailsText}");
    Console.WriteLine();

    var detailedProducts = new List<ProductDetails>();
    var sellerLegalInfoCache = new Dictionary<string, SellerLegalInfo>(StringComparer.OrdinalIgnoreCase);

    for (int i = 0; i < searchProducts.Count; i++)
    {
        var product = searchProducts[i];
        Console.WriteLine($"[{i + 1}/{searchProducts.Count}] {product.Title}");

        try
        {
            detailedProducts.Add(LoadProductDetails(driver, js, wait, product, sellerLegalInfoCache));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   {CouldNotProcessText}: {ex.Message}");
            detailedProducts.Add(ProductDetails.FromFailedProduct(product, ex.Message));
        }
    }

    var outputPath = CreateExcelReport(query, detailedProducts);

    Console.WriteLine();
    Console.WriteLine($"{ProcessedProductsText}: {detailedProducts.Count}");
    Console.WriteLine($"{ExcelSavedText}: {outputPath}");
}
catch (WebDriverTimeoutException)
{
    Console.WriteLine(WaitTimeoutText);
}
catch (Exception ex)
{
    Console.WriteLine($"{GenericErrorPrefix}: {ex.Message}");
}

static bool IsSearchResultsLoaded(IWebDriver drv)
{
    try
    {
        ClosePopupsIfAny(drv);
        return drv.FindElements(By.XPath(SearchResultsXPath)).Count > 0 &&
               drv.FindElements(By.XPath(ProductCardXPath)).Count > 0;
    }
    catch
    {
        return false;
    }
}

static string ReadSearchQuery()
{
    Console.Write(QueryPromptText);

    if (NativeConsole.TryReadUnicodeLine(out var query))
        return query.Trim();

    return Console.ReadLine()?.Trim() ?? string.Empty;
}

static List<ProductSummary> DeduplicateProducts(List<ProductSummary> rawProducts)
{
    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var result = new List<ProductSummary>();

    foreach (var product in rawProducts)
    {
        var dedupeUrl = GetDedupeUrl(product.Url);
        if (!string.IsNullOrWhiteSpace(dedupeUrl) && seen.Add(dedupeUrl))
        {
            result.Add(product);
        }
    }

    return result;
}

static ProductDetails LoadProductDetails(
    IWebDriver driver,
    IJavaScriptExecutor js,
    WebDriverWait wait,
    ProductSummary product,
    IDictionary<string, SellerLegalInfo> sellerLegalInfoCache)
{
    driver.Navigate().GoToUrl(product.Url);

    wait.Until(drv =>
    {
        try
        {
            ClosePopupsIfAny(drv);
            var page = drv.PageSource;

            return drv.FindElements(By.CssSelector("[data-auto='product-description']")).Count > 0 ||
                   drv.FindElements(By.CssSelector("[data-auto='specs-list-fullExtended']")).Count > 0 ||
                   drv.FindElements(By.CssSelector("[data-auto='shop-info-block']")).Count > 0 ||
                   page.Contains("articleNumberSpec", StringComparison.OrdinalIgnoreCase) ||
                   page.Contains("shopInfoBlock", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    });

    Thread.Sleep(1000);
    ClosePopupsIfAny(driver);

    var pageSource = driver.PageSource;
    var htmlDetails = ExtractProductDetailsFromHtml(pageSource);
    var jsDetails = ExtractProductDetailsViaJs(js);
    MergeProductDetails(htmlDetails, jsDetails);

    htmlDetails.Query = product.Query;
    htmlDetails.Title = product.Title;
    htmlDetails.Price = product.Price;
    htmlDetails.ProductUrl = product.Url;
    htmlDetails.ProductPageTitle = NormalizeText(driver.Title);
    htmlDetails.MarketArticle = NormalizeText(htmlDetails.MarketArticle);
    htmlDetails.Description = NormalizeText(htmlDetails.Description);
    htmlDetails.SellerName = NormalizeText(htmlDetails.SellerName);
    htmlDetails.SellerUrl = NormalizeUrl(htmlDetails.SellerUrl);

    if (!string.IsNullOrWhiteSpace(htmlDetails.SellerUrl))
    {
        var sellerLegalInfo = GetSellerLegalInfoFromStorePage(driver, js, wait, htmlDetails.SellerUrl, sellerLegalInfoCache);
        MergeSellerLegalData(htmlDetails, sellerLegalInfo);
    }

    if (!string.IsNullOrWhiteSpace(htmlDetails.MarketArticle) &&
        !htmlDetails.Specs.ContainsKey(MarketArticleLabel))
    {
        htmlDetails.Specs[MarketArticleLabel] = htmlDetails.MarketArticle;
    }

    return htmlDetails;
}

static void MergeProductDetails(ProductDetails target, ProductDetails source)
{
    if (string.IsNullOrWhiteSpace(target.MarketArticle))
        target.MarketArticle = source.MarketArticle;

    if (string.IsNullOrWhiteSpace(target.Description))
        target.Description = source.Description;

    if (string.IsNullOrWhiteSpace(target.SellerName))
        target.SellerName = source.SellerName;

    if (string.IsNullOrWhiteSpace(target.SellerUrl))
        target.SellerUrl = source.SellerUrl;

    foreach (var pair in source.Specs)
    {
        if (!target.Specs.ContainsKey(pair.Key))
        {
            target.Specs[pair.Key] = pair.Value;
        }
    }
}
static void MergeSellerLegalData(ProductDetails target, SellerLegalInfo source)
{
    if (string.IsNullOrWhiteSpace(target.SellerLegalName))
        target.SellerLegalName = source.LegalName;

    if (string.IsNullOrWhiteSpace(target.SellerInn))
        target.SellerInn = source.Inn;

    if (string.IsNullOrWhiteSpace(target.SellerOgrn))
        target.SellerOgrn = source.Ogrn;
}

static SellerLegalInfo GetSellerLegalInfoFromStorePage(
    IWebDriver driver,
    IJavaScriptExecutor js,
    WebDriverWait wait,
    string sellerUrl,
    IDictionary<string, SellerLegalInfo> cache)
{
    var normalizedSellerUrl = NormalizeUrl(sellerUrl);
    if (string.IsNullOrWhiteSpace(normalizedSellerUrl))
        return new SellerLegalInfo();

    if (cache.TryGetValue(normalizedSellerUrl, out var cachedInfo))
        return cachedInfo;

    var sellerInfo = new SellerLegalInfo();

    try
    {
        driver.Navigate().GoToUrl(normalizedSellerUrl);

        wait.Until(drv =>
        {
            try
            {
                ClosePopupsIfAny(drv);
                return ((IJavaScriptExecutor)drv).ExecuteScript("return document.readyState")?.ToString() == "complete";
            }
            catch
            {
                return false;
            }
        });

        Thread.Sleep(1000);
        ClosePopupsIfAny(driver);
        WaitForStoreInfoButton(driver, 7000);

        TryOpenSellerRatingPopup(driver, js);
        Thread.Sleep(700);
        TryExpandSellerLegalSection(driver, js);
        Thread.Sleep(700);

        sellerInfo = CaptureSellerLegalInfoFromCurrentPage(driver, js);
    }
    catch
    {
        sellerInfo = new SellerLegalInfo();
    }

    cache[normalizedSellerUrl] = sellerInfo;
    return sellerInfo;
}

static void TryOpenSellerRatingPopup(IWebDriver driver, IJavaScriptExecutor js)
{
    if (HasSellerLegalPopup(driver))
        return;

    var xpaths = new[]
    {
        "//button[@aria-label='Информация о магазине']",
        "//button[contains(@aria-label,'Информация о магазине')]",
        "//button[.//span[contains(normalize-space(),'оценок')]]",
        "//button[.//span[contains(normalize-space(),'подписчиков')]]"
    };

    foreach (var xpath in xpaths)
    {
        IReadOnlyCollection<IWebElement> elements;
        try
        {
            elements = driver.FindElements(By.XPath(xpath));
        }
        catch
        {
            continue;
        }

        foreach (var element in elements)
        {
            try
            {
                if (!IsElementDisplayed(element))
                    continue;

                if (!TryClickElement(driver, js, element))
                    continue;

                if (WaitForSellerLegalPopup(driver, 5000))
                    return;
            }
            catch
            {
            }
        }
    }

    try
    {
        var fallbackElements = driver.FindElements(By.CssSelector(SellerRatingBlockSelector));
        foreach (var element in fallbackElements)
        {
            try
            {
                if (!IsElementDisplayed(element))
                    continue;

                if (!TryClickElement(driver, js, element))
                    continue;

                if (WaitForSellerLegalPopup(driver, 5000))
                    return;
            }
            catch
            {
            }
        }
    }
    catch
    {
    }
}

static bool WaitForSellerLegalPopup(IWebDriver driver, int timeoutMs)
{
    var endTime = Environment.TickCount64 + timeoutMs;

    while (Environment.TickCount64 < endTime)
    {
        if (HasSellerLegalPopup(driver))
            return true;

        Thread.Sleep(200);
    }

    return HasSellerLegalPopup(driver);
}

static bool WaitForStoreInfoButton(IWebDriver driver, int timeoutMs)
{
    var endTime = Environment.TickCount64 + timeoutMs;

    while (Environment.TickCount64 < endTime)
    {
        try
        {
            var buttons = driver.FindElements(By.XPath("//button[@aria-label='Информация о магазине']"));
            if (buttons.Any(IsElementDisplayed))
                return true;

            buttons = driver.FindElements(By.XPath("//button[contains(@aria-label,'Информация о магазине')]"));
            if (buttons.Any(IsElementDisplayed))
                return true;
        }
        catch
        {
        }

        Thread.Sleep(200);
    }

    return false;
}

static bool HasSellerLegalPopup(IWebDriver driver)
{
    try
    {
        var accordionButtons = driver.FindElements(By.XPath(
            $"//button[normalize-space()='{SellerLegalInfoHeading}' or .//span[normalize-space()='{SellerLegalInfoHeading}']]"));

        if (accordionButtons.Any(IsElementDisplayed))
            return true;
    }
    catch
    {
    }

    try
    {
        var popups = driver.FindElements(By.CssSelector(SellerLegalPopupSelector));
        if (popups.Any(IsElementDisplayed))
        {
            return popups.Any(popup =>
            {
                try
                {
                    var text = NormalizeText(popup.Text);
                    return text.Contains(SellerLegalInfoHeading, StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return false;
                }
            });
        }
    }
    catch
    {
    }

    return false;
}

static void TryExpandSellerLegalSection(IWebDriver driver, IJavaScriptExecutor js)
{
    try
    {
        var accordionContainers = driver.FindElements(By.CssSelector(SellerLegalAccordionSelector));
        foreach (var accordionContainer in accordionContainers)
        {
            try
            {
                var button = accordionContainer.FindElements(By.CssSelector("button, [role='button']")).FirstOrDefault();
                if (button is null)
                    continue;

                if (TryClickElement(driver, js, button))
                {
                    Thread.Sleep(500);
                    return;
                }
            }
            catch
            {
            }
        }
    }
    catch
    {
    }

    var xpaths = new[]
    {
        $"//button[normalize-space()='{SellerLegalInfoHeading}']",
        $"//button[.//span[normalize-space()='{SellerLegalInfoHeading}']]",
        $"//*[@role='button'][.//span[normalize-space()='{SellerLegalInfoHeading}']]",
        $"//button[contains(normalize-space(),'{SellerLegalInfoHeading}')]"
    };

    foreach (var xpath in xpaths)
    {
        IReadOnlyCollection<IWebElement> buttons;
        try
        {
            buttons = driver.FindElements(By.XPath(xpath));
        }
        catch
        {
            continue;
        }

        foreach (var button in buttons)
        {
            try
            {
                if (TryClickElement(driver, js, button))
                {
                    Thread.Sleep(500);
                    return;
                }
            }
            catch
            {
            }
        }
    }
}

static bool TryClickElement(IWebDriver driver, IJavaScriptExecutor js, IWebElement element)
{
    if (!IsElementDisplayed(element))
        return false;

    try
    {
        js.ExecuteScript(
            """
            const el = arguments[0];
            el.scrollIntoView({ block: 'center', inline: 'center' });
            """,
            element);
        Thread.Sleep(150);
    }
    catch
    {
    }

    try
    {
        element.Click();
        Thread.Sleep(250);
        return true;
    }
    catch
    {
    }

    try
    {
        js.ExecuteScript(
            """
            const el = arguments[0];
            const rect = el.getBoundingClientRect();
            const x = rect.left + rect.width / 2;
            const y = rect.top + rect.height / 2;

            for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                el.dispatchEvent(new MouseEvent(type, {
                    view: window,
                    bubbles: true,
                    cancelable: true,
                    clientX: x,
                    clientY: y,
                    buttons: 1
                }));
            }
            """,
            element);
        Thread.Sleep(350);
        return true;
    }
    catch
    {
    }

    try
    {
        js.ExecuteScript("arguments[0].click();", element);
        Thread.Sleep(250);
        return true;
    }
    catch
    {
        return false;
    }
}

static bool IsElementDisplayed(IWebElement element)
{
    try
    {
        return element.Displayed && element.Enabled;
    }
    catch
    {
        return false;
    }
}

static SellerLegalInfo CaptureSellerLegalInfoFromCurrentPage(IWebDriver driver, IJavaScriptExecutor js)
{
    var popupHtml = ExtractSellerPopupHtml(js);
    var popupText = ExtractSellerPopupText(driver);
    var pageSource = driver.PageSource;

    var sellerInfo = ExtractSellerLegalInfoFromPopupMarkup(popupHtml);
    if (!sellerInfo.HasAnyValue)
        sellerInfo = ExtractSellerLegalInfoFromPopupMarkup(pageSource);

    if (!sellerInfo.HasAnyValue)
        sellerInfo = ExtractSellerLegalInfoFromPopupText(popupText);

    if (!sellerInfo.HasAnyValue)
        sellerInfo = ExtractSellerLegalInfoFromPopupText(GetBodyVisibleText(js));

    sellerInfo.LegalName = NormalizeText(sellerInfo.LegalName);
    sellerInfo.Inn = NormalizeDigitList(sellerInfo.Inn, 10, 12);
    sellerInfo.Ogrn = NormalizeDigitList(sellerInfo.Ogrn, 13, 15);

    return sellerInfo;
}

static string ExtractSellerPopupHtml(IJavaScriptExecutor js)
{
    try
    {
        return js.ExecuteScript("const popup = document.querySelector(\"[id='/content/popup']\"); return popup ? popup.outerHTML : '';")?.ToString() ?? string.Empty;
    }
    catch
    {
        return string.Empty;
    }
}

static string ExtractSellerPopupText(IWebDriver driver)
{
    try
    {
        var popup = driver.FindElements(By.CssSelector(SellerLegalPopupSelector)).FirstOrDefault();
        return popup is null ? string.Empty : NormalizeMultilineTextActive(popup.Text);
    }
    catch
    {
        return string.Empty;
    }
}

static string GetBodyVisibleText(IJavaScriptExecutor js)
{
    try
    {
        return NormalizeMultilineTextActive(js.ExecuteScript("return document.body ? document.body.innerText : '';")?.ToString());
    }
    catch
    {
        return string.Empty;
    }
}

static SellerLegalInfo ExtractSellerLegalInfoFromPopupMarkup(string html)
{
    if (string.IsNullOrWhiteSpace(html))
        return new SellerLegalInfo();

    var matches = Regex.Matches(
        html,
        @"<noframes[^>]*data-apiary=""patch""[^>]*>(?<json>\{.*?\})</noframes>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    foreach (Match match in matches)
    {
        var info = ParseSellerLegalInfoFromPatchJson(WebUtility.HtmlDecode(match.Groups["json"].Value));
        if (info.HasAnyValue)
            return info;
    }

    return new SellerLegalInfo();
}

static SellerLegalInfo ParseSellerLegalInfoFromPatchJson(string patchJson)
{
    if (string.IsNullOrWhiteSpace(patchJson))
        return new SellerLegalInfo();

    try
    {
        using var document = JsonDocument.Parse(patchJson);
        if (!TryFindJsonPropertyRecursive(document.RootElement, "shopsJurData", out var shopsJurData) ||
            shopsJurData.ValueKind != JsonValueKind.Array)
        {
            return new SellerLegalInfo();
        }

        var legalNames = new List<string>();
        var inns = new List<string>();
        var ogrns = new List<string>();

        foreach (var item in shopsJurData.EnumerateArray())
        {
            AddUniqueTextValue(legalNames, GetJsonText(item, "legalName"));
            AddUniqueTextValue(inns, GetNestedJsonText(item, "innEntity", "value"));

            var ogrn = GetNestedJsonText(item, "ogrnEntity", "value");
            if (string.IsNullOrWhiteSpace(ogrn))
                ogrn = GetNestedJsonText(item, "registrationNumberEntity", "value");

            AddUniqueTextValue(ogrns, ogrn);
        }

        return new SellerLegalInfo
        {
            LegalName = string.Join(" | ", legalNames),
            Inn = string.Join(" | ", inns),
            Ogrn = string.Join(" | ", ogrns)
        };
    }
    catch
    {
        return new SellerLegalInfo();
    }
}

static bool TryFindJsonPropertyRecursive(JsonElement element, string propertyName, out JsonElement value)
{
    if (element.ValueKind == JsonValueKind.Object)
    {
        foreach (var property in element.EnumerateObject())
        {
            if (property.NameEquals(propertyName))
            {
                value = property.Value;
                return true;
            }

            if (TryFindJsonPropertyRecursive(property.Value, propertyName, out value))
                return true;
        }
    }
    else if (element.ValueKind == JsonValueKind.Array)
    {
        foreach (var item in element.EnumerateArray())
        {
            if (TryFindJsonPropertyRecursive(item, propertyName, out value))
                return true;
        }
    }

    value = default;
    return false;
}

static string GetJsonText(JsonElement element, string propertyName)
{
    if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out var property))
        return string.Empty;

    return property.ValueKind == JsonValueKind.String
        ? NormalizeText(property.GetString())
        : NormalizeText(property.ToString());
}

static string GetNestedJsonText(JsonElement element, string outerPropertyName, string innerPropertyName)
{
    if (element.ValueKind != JsonValueKind.Object ||
        !element.TryGetProperty(outerPropertyName, out var outer) ||
        outer.ValueKind != JsonValueKind.Object ||
        !outer.TryGetProperty(innerPropertyName, out var inner))
    {
        return string.Empty;
    }

    return inner.ValueKind == JsonValueKind.String
        ? NormalizeText(inner.GetString())
        : NormalizeText(inner.ToString());
}

static SellerLegalInfo ExtractSellerLegalInfoFromPopupText(string text)
{
    if (string.IsNullOrWhiteSpace(text))
        return new SellerLegalInfo();

    var legalNames = new List<string>();
    var inns = new List<string>();
    var ogrns = new List<string>();

    var lines = NormalizeMultilineTextActive(text)
        .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Select(NormalizeText)
        .Where(line => !string.IsNullOrWhiteSpace(line))
        .ToList();

    foreach (var line in lines)
    {
        var hasInn = line.Contains(InnLabel, StringComparison.OrdinalIgnoreCase);
        var hasOgrn = line.Contains(OgrnLabel, StringComparison.OrdinalIgnoreCase) ||
                      line.Contains(OgrnipLabel, StringComparison.OrdinalIgnoreCase);

        if (!hasInn && !hasOgrn)
            continue;

        AddUniqueTextValue(legalNames, ExtractSellerLegalNameFromLine(line));
        AddUniqueTextValue(inns, ExtractRegexValue(line, @"\bИНН\b[^\d]{0,20}(?<value>\d{10}|\d{12})\b"));
        AddUniqueTextValue(ogrns, ExtractRegexValue(line, @"\b(?:ОГРН|ОГРНИП)\b[^\d]{0,20}(?<value>\d{13}|\d{15})\b"));
    }

    return new SellerLegalInfo
    {
        LegalName = string.Join(" | ", legalNames),
        Inn = string.Join(" | ", inns),
        Ogrn = string.Join(" | ", ogrns)
    };
}

static string ExtractSellerLegalNameFromLine(string line)
{
    return ExtractRegexValue(
        line,
        @"(?<value>(?:ИП|ООО|ОАО|ЗАО|АО|ПАО|НАО|Самозанятый)\s+.+?)(?=\s+(?:Юридический адрес|ОГРН|ОГРНИП|ИНН)\b|$)");
}

static string ExtractRegexValue(string text, string pattern)
{
    if (string.IsNullOrWhiteSpace(text))
        return string.Empty;

    var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
    return match.Success ? NormalizeText(match.Groups["value"].Value) : string.Empty;
}

static void AddUniqueTextValue(ICollection<string> values, string candidate)
{
    var normalized = NormalizeText(candidate);
    if (string.IsNullOrWhiteSpace(normalized))
        return;

    if (values.Any(existing => existing.Equals(normalized, StringComparison.OrdinalIgnoreCase)))
        return;

    values.Add(normalized);
}

static string NormalizeDigitList(string value, params int[] allowedLengths)
{
    if (string.IsNullOrWhiteSpace(value))
        return string.Empty;

    var matches = Regex.Matches(value, @"\d+")
        .Select(match => match.Value)
        .Where(digits => allowedLengths.Length == 0 || allowedLengths.Contains(digits.Length))
        .Distinct(StringComparer.Ordinal)
        .ToList();

    return string.Join(" | ", matches);
}

static string NormalizeMultilineTextActive(string? value)
{
    if (string.IsNullOrWhiteSpace(value))
        return string.Empty;

    var normalized = WebUtility.HtmlDecode(DecodeHtmlUnicode(value))
        .Replace('\u2006', ' ')
        .Replace('\u202F', ' ')
        .Replace('\u00A0', ' ')
        .Replace("\r\n", "\n")
        .Replace('\r', '\n');

    normalized = Regex.Replace(normalized, @"[ \t\f\v]+", " ");
    normalized = Regex.Replace(normalized, @"\n{2,}", "\n");

    return RemoveInvalidXmlChars(normalized.Trim());
}

static ProductDetails ExtractProductDetailsFromHtml(string html)
{
    var details = new ProductDetails();

    details.Description = ExtractDescriptionFromHtml(html);
    details.MarketArticle = ExtractMarketArticleFromHtml(html);

    foreach (var pair in ExtractSpecsFromHtml(html))
    {
        details.Specs[pair.Key] = pair.Value;
    }

    if (!string.IsNullOrWhiteSpace(details.MarketArticle) &&
        !details.Specs.ContainsKey(MarketArticleLabel))
    {
        details.Specs[MarketArticleLabel] = details.MarketArticle;
    }

    var seller = ExtractSellerFromHtml(html);
    details.SellerName = seller.Name;
    details.SellerUrl = seller.Url;

    return details;
}

static string ExtractDescriptionFromHtml(string html)
{
    if (string.IsNullOrWhiteSpace(html))
        return string.Empty;

    var match = Regex.Match(
        html,
        @"data-auto=""product-description"".*?<div class=""ds-text[^""]*"">(?<value>.*?)</div>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    if (!match.Success)
        return string.Empty;

    return NormalizeHtmlFragment(match.Groups["value"].Value);
}

static Dictionary<string, string> ExtractSpecsFromHtml(string html)
{
    var specs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    if (string.IsNullOrWhiteSpace(html))
        return specs;

    var nameMatches = Regex.Matches(
        html,
        @"<span\s+data-auto=""product-spec""[^>]*>(?<name>.*?)</span>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    foreach (Match nameMatch in nameMatches)
    {
        var name = NormalizeHtmlFragment(nameMatch.Groups["name"].Value);
        if (string.IsNullOrWhiteSpace(name))
            continue;

        var segmentLength = Math.Min(1800, html.Length - nameMatch.Index);
        var segment = html.Substring(nameMatch.Index, segmentLength);

        var valueMatch = Regex.Match(
            segment,
            @"<div class=""eXP5k"".*?<span>(?<value>.*?)</span>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        if (!valueMatch.Success)
        {
            valueMatch = Regex.Match(
                segment,
                @"<div class=""eXP5k"".*?<div class=""ds-text[^""]*"">(?<value>.*?)</div>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
        }

        if (!valueMatch.Success)
            continue;

        var value = NormalizeHtmlFragment(valueMatch.Groups["value"].Value);
        if (string.IsNullOrWhiteSpace(value))
            continue;

        specs[name] = value;
    }

    return specs;
}

static string ExtractMarketArticleFromHtml(string html)
{
    if (string.IsNullOrWhiteSpace(html))
        return string.Empty;

    var widgetMatch = Regex.Match(
        html,
        @"articleNumberSpec.*?<span>(?<value>[^<]+)</span>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    if (widgetMatch.Success)
        return NormalizeHtmlFragment(widgetMatch.Groups["value"].Value);

    var jsonMatch = Regex.Match(
        html,
        @"""articleNumberSpec"".{0,800}?""value"":""(?<value>[^""]+)""",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    return jsonMatch.Success ? NormalizeText(jsonMatch.Groups["value"].Value) : string.Empty;
}

static SellerInfo ExtractSellerFromHtml(string html)
{
    if (string.IsNullOrWhiteSpace(html))
        return new SellerInfo();

    var businessUrl = ExtractBusinessSellerUrl(html);

    var jsonMatch = Regex.Match(
        html,
        @"""shopInfoBlock"":\{.*?""name"":""(?<name>[^""]+)""[^}]*?""searchPageLink"":""(?<url>[^""]+)""",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    if (jsonMatch.Success)
    {
        return new SellerInfo
        {
            Name = NormalizeText(jsonMatch.Groups["name"].Value),
            Url = string.IsNullOrWhiteSpace(businessUrl)
                ? NormalizeUrl(jsonMatch.Groups["url"].Value)
                : businessUrl
        };
    }

    var domMatch = Regex.Match(
        html,
        @"data-auto=""shop-info-block"".*?<a href=""(?<url>[^""]+)""[^>]*>.*?data-auto=""shop-info-title""[^>]*>.*?<span[^>]*>(?<name>.*?)</span>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    return new SellerInfo
    {
        Name = domMatch.Success ? NormalizeHtmlFragment(domMatch.Groups["name"].Value) : string.Empty,
        Url = !string.IsNullOrWhiteSpace(businessUrl)
            ? businessUrl
            : domMatch.Success ? NormalizeUrl(domMatch.Groups["url"].Value) : string.Empty
    };
}

static string ExtractBusinessSellerUrl(string html)
{
    if (string.IsNullOrWhiteSpace(html))
        return string.Empty;

    var patterns = new[]
    {
        @"(?<url>https?:\\?/\\?/market\.yandex\.ru\\?/business--[^""'\\\s<]+/\d+(?:\?[^""'\\<\s]*)?)",
        @"(?<url>/business--[^""'\\\s<]+/\d+(?:\?[^""'\\<\s]*)?)"
    };

    foreach (var pattern in patterns)
    {
        var match = Regex.Match(html, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
        if (!match.Success)
            continue;

        var normalized = NormalizeUrl(match.Groups["url"].Value);
        if (!string.IsNullOrWhiteSpace(normalized))
            return normalized;
    }

    return string.Empty;
}

static void ScrollUntilAllProductsLoaded(IWebDriver driver, IJavaScriptExecutor js)
{
    int stableIterations = 0;
    long previousCount = -1;
    const int maxStableIterations = 5;

    while (stableIterations < maxStableIterations)
    {
        try
        {
            ClosePopupsIfAny(driver);

            long countBefore = GetCardCountViaJs(js);
            js.ExecuteScript(ScrollJs);

            Thread.Sleep(2000);
            ClosePopupsIfAny(driver);

            long countAfterFirst = GetCardCountViaJs(js);

            if (countAfterFirst == countBefore)
            {
                Thread.Sleep(1500);
                long countAfterSecond = GetCardCountViaJs(js);

                if (countAfterSecond == countBefore && countAfterSecond == previousCount)
                {
                    stableIterations++;
                }
                else
                {
                    stableIterations = 0;
                }

                previousCount = countAfterSecond;
            }
            else
            {
                stableIterations = 0;
                previousCount = countAfterFirst;
            }
        }
        catch
        {
            Thread.Sleep(1000);
            stableIterations++;
        }
    }

    Thread.Sleep(1000);
}

static long GetCardCountViaJs(IJavaScriptExecutor js)
{
    return Convert.ToInt64(js.ExecuteScript(CountCardsJs));
}

static List<ProductSummary> ExtractAllProductsViaJs(IJavaScriptExecutor js)
{
    var products = new List<ProductSummary>();

    var result = js.ExecuteScript(ExtractProductsJs);
    if (result is not ReadOnlyCollection<object> items)
        return products;

    foreach (var item in items)
    {
        if (item is not Dictionary<string, object> dict)
            continue;

        var title = dict.TryGetValue("title", out var titleValue) ? titleValue?.ToString() ?? string.Empty : string.Empty;
        var url = dict.TryGetValue("url", out var urlValue) ? urlValue?.ToString() ?? string.Empty : string.Empty;
        var rawPrice = dict.TryGetValue("price", out var priceValue) ? priceValue?.ToString() ?? string.Empty : string.Empty;

        title = NormalizeText(title);
        url = NormalizeUrl(url);
        rawPrice = NormalizeText(rawPrice);

        if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(url))
            continue;

        products.Add(new ProductSummary
        {
            Title = title,
            Price = string.IsNullOrWhiteSpace(rawPrice) ? NoPriceText : NormalizePrice(rawPrice),
            Url = url
        });
    }

    return products;
}

static ProductDetails ExtractProductDetailsViaJs(IJavaScriptExecutor js)
{
    try
    {
        var details = new ProductDetails();
        var result = js.ExecuteScript(ExtractProductDetailsJs);

        if (result is not Dictionary<string, object> dict)
            return details;

        details.Description = dict.TryGetValue("description", out var descriptionValue)
            ? NormalizeText(descriptionValue?.ToString())
            : string.Empty;

        details.SellerName = dict.TryGetValue("sellerName", out var sellerNameValue)
            ? NormalizeText(sellerNameValue?.ToString())
            : string.Empty;

        details.SellerUrl = dict.TryGetValue("sellerUrl", out var sellerUrlValue)
            ? NormalizeUrl(sellerUrlValue?.ToString() ?? string.Empty)
            : string.Empty;

        if (dict.TryGetValue("specs", out var specsValue) && specsValue is ReadOnlyCollection<object> specsItems)
        {
            foreach (var item in specsItems)
            {
                if (item is not Dictionary<string, object> specDict)
                    continue;

                var name = specDict.TryGetValue("name", out var nameValue) ? NormalizeText(nameValue?.ToString()) : string.Empty;
                var value = specDict.TryGetValue("value", out var valueValue) ? NormalizeText(valueValue?.ToString()) : string.Empty;

                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value))
                    continue;

                details.Specs[name] = value;

                if (name.Equals(MarketArticleLabel, StringComparison.OrdinalIgnoreCase))
                {
                    details.MarketArticle = value;
                }
            }
        }

        return details;
    }
    catch
    {
        return new ProductDetails();
    }
}

static void ClosePopupsIfAny(IWebDriver driver)
{
    try
    {
        var buttons = driver.FindElements(By.XPath(PopupButtonXPath));
        foreach (var button in buttons)
        {
            try
            {
                button.Click();
                Thread.Sleep(500);
                break;
            }
            catch
            {
            }
        }
    }
    catch
    {
    }
}

static string CreateExcelReport(string query, IReadOnlyCollection<ProductDetails> products)
{
    var outputDirectory = Path.Combine(AppContext.BaseDirectory, "outputs");
    Directory.CreateDirectory(outputDirectory);

    var fileName = $"market_{BuildSafeFileName(query)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
    var outputPath = Path.Combine(outputDirectory, fileName);
    var tempPath = Path.Combine(outputDirectory, $"{Guid.NewGuid():N}.tmp.xlsx");

    var specColumns = products
        .SelectMany(product => product.Specs.Keys)
        .Where(name => !name.Equals(MarketArticleLabel, StringComparison.OrdinalIgnoreCase))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
        .ToList();
    var sellerLegalColumnGroupCount = Math.Max(
        1,
        products.Select(GetSellerLegalColumnGroupCount).DefaultIfEmpty(0).Max());

    var headers = new List<string>
    {
        "\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u0438\u0437 \u043f\u043e\u0438\u0441\u043a\u0430",
        "\u0426\u0435\u043d\u0430",
        "\u0421\u0441\u044b\u043b\u043a\u0430 \u043d\u0430 \u0442\u043e\u0432\u0430\u0440",
        MarketArticleLabel,
        "\u041e\u043f\u0438\u0441\u0430\u043d\u0438\u0435",
        "\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u043c\u0430\u0433\u0430\u0437\u0438\u043d\u0430",
        "\u0421\u0441\u044b\u043b\u043a\u0430 \u043d\u0430 \u043c\u0430\u0433\u0430\u0437\u0438\u043d",
        "\u041f\u043e\u0438\u0441\u043a\u043e\u0432\u044b\u0439 \u0437\u0430\u043f\u0440\u043e\u0441",
        "\u0417\u0430\u0433\u043e\u043b\u043e\u0432\u043e\u043a \u0441\u0442\u0440\u0430\u043d\u0438\u0446\u044b \u0442\u043e\u0432\u0430\u0440\u0430",
        "\u041e\u0448\u0438\u0431\u043a\u0430 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438"
    };

    headers.InsertRange(7, BuildSellerLegalHeaders(sellerLegalColumnGroupCount));

    headers.AddRange(specColumns);

    var rows = products
        .Select(product =>
        {
            var rowValues = new List<string>
            {
                SanitizeForExcel(product.Title),
                SanitizeForExcel(product.Price),
                SanitizeForExcel(product.ProductUrl),
                SanitizeForExcel(product.MarketArticle),
                SanitizeForExcel(product.Description),
                SanitizeForExcel(product.SellerName),
                SanitizeForExcel(product.SellerUrl),
                SanitizeForExcel(product.Query),
                SanitizeForExcel(product.ProductPageTitle),
                SanitizeForExcel(product.ErrorMessage)
            };

            rowValues.InsertRange(7, BuildSellerLegalRowValues(product, sellerLegalColumnGroupCount));

            for (int i = 0; i < specColumns.Count; i++)
            {
                var specName = specColumns[i];
                rowValues.Add(product.Specs.TryGetValue(specName, out var value)
                    ? SanitizeForExcel(value)
                    : string.Empty);
            }

            return (IReadOnlyList<string>)rowValues;
        })
        .ToList();

    WriteExcelWorkbook(tempPath, headers, rows);

    if (File.Exists(outputPath))
        File.Delete(outputPath);

    File.Move(tempPath, outputPath);
    return outputPath;
}

static int GetSellerLegalColumnGroupCount(ProductDetails product)
{
    return new[]
    {
        SplitSellerLegalValues(product.SellerLegalName).Count,
        SplitSellerLegalValues(product.SellerInn).Count,
        SplitSellerLegalValues(product.SellerOgrn).Count
    }.Max();
}

static List<string> BuildSellerLegalHeaders(int sellerLegalColumnGroupCount)
{
    var headers = new List<string>(sellerLegalColumnGroupCount * 3);

    for (var index = 1; index <= sellerLegalColumnGroupCount; index++)
    {
        headers.Add($"\u042e\u0440 \u043b\u0438\u0446\u043e {index}");
        headers.Add($"{InnLabel} {index}");
        headers.Add($"\u041e\u0413\u0420\u041d/\u041e\u0413\u0420\u041d\u0418\u041f {index}");
    }

    return headers;
}

static List<string> BuildSellerLegalRowValues(ProductDetails product, int sellerLegalColumnGroupCount)
{
    var legalNames = SplitSellerLegalValues(product.SellerLegalName);
    var inns = SplitSellerLegalValues(product.SellerInn);
    var ogrns = SplitSellerLegalValues(product.SellerOgrn);

    var rowValues = new List<string>(sellerLegalColumnGroupCount * 3);

    for (var index = 0; index < sellerLegalColumnGroupCount; index++)
    {
        rowValues.Add(index < legalNames.Count ? SanitizeForExcel(legalNames[index]) : string.Empty);
        rowValues.Add(index < inns.Count ? SanitizeForExcel(inns[index]) : string.Empty);
        rowValues.Add(index < ogrns.Count ? SanitizeForExcel(ogrns[index]) : string.Empty);
    }

    return rowValues;
}

static List<string> SplitSellerLegalValues(string value)
{
    if (string.IsNullOrWhiteSpace(value))
        return new List<string>();

    return value
        .Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Where(item => !string.IsNullOrWhiteSpace(item))
        .ToList();
}

static string NormalizeUrl(string url)
{
    if (string.IsNullOrWhiteSpace(url))
        return string.Empty;

    var decoded = WebUtility.HtmlDecode(DecodeHtmlUnicode(url))
        .Replace("\\/", "/")
        .Trim();

    if (Uri.TryCreate(decoded, UriKind.Absolute, out var absoluteUri))
        return absoluteUri.ToString();

    if (Uri.TryCreate(new Uri("https://market.yandex.ru"), decoded, out var relativeUri))
        return relativeUri.ToString();

    return decoded;
}

static void WriteExcelWorkbook(
    string outputPath,
    IReadOnlyList<string> headers,
    IReadOnlyList<IReadOnlyList<string>> rows)
{
    using var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook);

    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = CreateStylesheet();
    stylesPart.Stylesheet.Save();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData();

    var headerRow = new Row { RowIndex = 1U };
    for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++)
    {
        var cellReference = GetCellReference(columnIndex + 1, 1U);
        headerRow.Append(CreateTextCell(headers[columnIndex], 1U, cellReference));
    }

    sheetData.Append(headerRow);

    var hyperlinks = new Hyperlinks();

    for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
    {
        var values = rows[rowIndex];
        var rowNumber = (uint)(rowIndex + 2);
        var row = new Row { RowIndex = rowNumber };

        for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++)
        {
            var value = columnIndex < values.Count ? values[columnIndex] : string.Empty;
            var cellReference = GetCellReference(columnIndex + 1, rowNumber);

            row.Append(CreateTextCell(value, 2U, cellReference));

            if ((columnIndex == 2 || columnIndex == 6) &&
                Uri.TryCreate(value, UriKind.Absolute, out var hyperlinkUri))
            {
                var relationship = worksheetPart.AddHyperlinkRelationship(hyperlinkUri, true);
                hyperlinks.Append(new Hyperlink
                {
                    Reference = cellReference,
                    Id = relationship.Id
                });
            }
        }

        sheetData.Append(row);
    }

    var worksheet = new Worksheet();
    worksheet.Append(new SheetViews(
        new SheetView
        {
            WorkbookViewId = 0U,
            Pane = new Pane
            {
                VerticalSplit = 1D,
                TopLeftCell = "A2",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            }
        }));

    var columns = BuildWorksheetColumns(headers, rows);
    if (columns.HasChildren)
    {
        worksheet.Append(columns);
    }

    worksheet.Append(sheetData);

    if (headers.Count > 0)
    {
        worksheet.Append(new AutoFilter
        {
            Reference = $"A1:{GetExcelColumnName(headers.Count)}{rows.Count + 1}"
        });
    }

    if (hyperlinks.HasChildren)
    {
        worksheet.Append(hyperlinks);
    }

    worksheetPart.Worksheet = worksheet;
    worksheetPart.Worksheet.Save();

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet
    {
        Id = workbookPart.GetIdOfPart(worksheetPart),
        SheetId = 1U,
        Name = "\u0422\u043e\u0432\u0430\u0440\u044b"
    });

    workbookPart.Workbook.Save();
}

static Columns BuildWorksheetColumns(
    IReadOnlyList<string> headers,
    IReadOnlyList<IReadOnlyList<string>> rows)
{
    var columns = new Columns();

    for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++)
    {
        var maxLength = headers[columnIndex].Length;

        foreach (var row in rows)
        {
            if (columnIndex >= row.Count)
                continue;

            maxLength = Math.Max(maxLength, MeasureCellLength(row[columnIndex]));
        }

        columns.Append(new Column
        {
            Min = (uint)(columnIndex + 1),
            Max = (uint)(columnIndex + 1),
            Width = GetPreferredColumnWidth(columnIndex, maxLength),
            CustomWidth = true
        });
    }

    return columns;
}

static int MeasureCellLength(string? value)
{
    if (string.IsNullOrWhiteSpace(value))
        return 0;

    return value
        .Split('\n', StringSplitOptions.TrimEntries)
        .Select(part => Math.Min(part.Length, 80))
        .DefaultIfEmpty(0)
        .Max();
}

static double GetPreferredColumnWidth(int columnIndex, int maxLength)
{
    return columnIndex switch
    {
        0 => Math.Clamp(maxLength + 4, 28, 45),
        1 => Math.Clamp(maxLength + 2, 12, 18),
        2 => Math.Clamp(maxLength + 4, 30, 55),
        3 => Math.Clamp(maxLength + 2, 16, 24),
        4 => Math.Clamp(maxLength + 4, 30, 60),
        5 => Math.Clamp(maxLength + 3, 18, 32),
        6 => Math.Clamp(maxLength + 4, 28, 55),
        7 => Math.Clamp(maxLength + 2, 18, 28),
        8 => Math.Clamp(maxLength + 4, 24, 45),
        9 => Math.Clamp(maxLength + 4, 24, 45),
        _ => Math.Clamp(maxLength + 3, 18, 30)
    };
}

static Cell CreateTextCell(string? value, uint styleIndex, string cellReference)
{
    return new Cell
    {
        CellReference = cellReference,
        StyleIndex = styleIndex,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(
            new Text(SanitizeForExcel(value))
            {
                Space = SpaceProcessingModeValues.Preserve
            })
    };
}

static Stylesheet CreateStylesheet()
{
    return new Stylesheet(
        new Fonts(
            new Font(
                new FontSize { Val = 11D },
                new FontName { Val = "Calibri" }),
            new Font(
                new Bold(),
                new FontSize { Val = 11D },
                new FontName { Val = "Calibri" }))
        {
            Count = 2U
        },
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
            new Fill(
                new PatternFill(
                    new ForegroundColor { Rgb = "FFF2F2F2" })
                {
                    PatternType = PatternValues.Solid
                }))
        {
            Count = 3U
        },
        new Borders(
            new Border(
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(),
                new DiagonalBorder()))
        {
            Count = 1U
        },
        new CellStyleFormats(new CellFormat())
        {
            Count = 1U
        },
        new CellFormats(
            new CellFormat
            {
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            },
            new CellFormat
            {
                FontId = 1U,
                FillId = 2U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyFont = true,
                ApplyFill = true,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Top,
                    WrapText = true
                }
            },
            new CellFormat
            {
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Top,
                    WrapText = true
                }
            })
        {
            Count = 3U
        },
        new CellStyles(
            new CellStyle
            {
                Name = "Normal",
                FormatId = 0U,
                BuiltinId = 0U
            })
        {
            Count = 1U
        },
        new DifferentialFormats { Count = 0U },
        new TableStyles
        {
            Count = 0U,
            DefaultTableStyle = "TableStyleMedium2",
            DefaultPivotStyle = "PivotStyleLight16"
        });
}

static string GetCellReference(int columnNumber, uint rowNumber)
{
    return $"{GetExcelColumnName(columnNumber)}{rowNumber}";
}

static string GetExcelColumnName(int columnNumber)
{
    if (columnNumber <= 0)
        throw new ArgumentOutOfRangeException(nameof(columnNumber));

    var builder = new StringBuilder();

    while (columnNumber > 0)
    {
        columnNumber--;
        builder.Insert(0, (char)('A' + (columnNumber % 26)));
        columnNumber /= 26;
    }

    return builder.ToString();
}

static string GetDedupeUrl(string url)
{
    if (string.IsNullOrWhiteSpace(url))
        return string.Empty;

    if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        return url;

    return $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}".TrimEnd('/');
}

static string NormalizePrice(string rawPrice)
{
    var normalized = NormalizeText(rawPrice);
    return normalized.Contains('\u20BD', StringComparison.Ordinal) ? normalized : $"{normalized} \u20BD";
}

static string NormalizeText(string? value)
{
    if (string.IsNullOrWhiteSpace(value))
        return string.Empty;

    var normalized = WebUtility.HtmlDecode(DecodeHtmlUnicode(value))
        .Replace('\u2006', ' ')
        .Replace('\u202F', ' ')
        .Replace('\u00A0', ' ')
        .Replace('\r', ' ')
        .Replace('\n', ' ')
        .Trim();

    normalized = Regex.Replace(normalized, @"\s+", " ");
    return RemoveInvalidXmlChars(normalized);
}

static string NormalizeHtmlFragment(string? html)
{
    if (string.IsNullOrWhiteSpace(html))
        return string.Empty;

    var text = Regex.Replace(html, @"<br\s*/?>", " ", RegexOptions.IgnoreCase);
    text = Regex.Replace(text, @"</(div|p|li|h\d|span)>", " ", RegexOptions.IgnoreCase);
    text = Regex.Replace(text, @"<[^>]+>", " ");

    return NormalizeText(text);
}

static string DecodeHtmlUnicode(string value)
{
    return Regex.Replace(
        value,
        @"\\u(?<code>[0-9a-fA-F]{4})",
        match => ((char)Convert.ToInt32(match.Groups["code"].Value, 16)).ToString());
}

static string BuildSafeFileName(string query)
{
    var normalized = RemoveDiacritics(query);
    var safe = Regex.Replace(normalized, @"[^A-Za-z0-9_-]+", "_").Trim('_');

    return string.IsNullOrWhiteSpace(safe) ? "report" : safe[..Math.Min(safe.Length, 50)];
}

static string SanitizeForExcel(string? value)
{
    return RemoveInvalidXmlChars(NormalizeText(value));
}

static string RemoveInvalidXmlChars(string? value)
{
    if (string.IsNullOrEmpty(value))
        return string.Empty;

    var builder = new StringBuilder(value.Length);
    foreach (var ch in value)
    {
        if (ch == 0x9 || ch == 0xA || ch == 0xD ||
            (ch >= 0x20 && ch <= 0xD7FF) ||
            (ch >= 0xE000 && ch <= 0xFFFD))
        {
            builder.Append(ch);
        }
    }

    return builder.ToString();
}

static string RemoveDiacritics(string value)
{
    var normalized = value.Normalize(NormalizationForm.FormD);
    var builder = new StringBuilder(normalized.Length);

    foreach (var ch in normalized)
    {
        var category = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch);
        if (category != System.Globalization.UnicodeCategory.NonSpacingMark)
            builder.Append(ch);
    }

    return builder.ToString().Normalize(NormalizationForm.FormC);
}

static class NativeConsole
{
    const int StdInputHandle = -10;
    const uint FileTypeChar = 0x0002;

    public static bool TryReadUnicodeLine(out string value)
    {
        value = string.Empty;

        if (!OperatingSystem.IsWindows() || Console.IsInputRedirected)
            return false;

        var stdIn = GetStdHandle(StdInputHandle);
        if (stdIn == IntPtr.Zero || stdIn == new IntPtr(-1))
            return false;

        if (GetFileType(stdIn) != FileTypeChar)
            return false;

        StringBuilder buffer = new(2048);
        if (!ReadConsoleW(stdIn, buffer, (uint)buffer.Capacity, out var charsRead, IntPtr.Zero) || charsRead == 0)
            return false;

        value = buffer
            .ToString(0, (int)charsRead)
            .TrimEnd('\r', '\n');

        return true;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern uint GetFileType(IntPtr hFile);

    [DllImport("kernel32.dll", EntryPoint = "ReadConsoleW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool ReadConsoleW(
        IntPtr hConsoleInput,
        StringBuilder lpBuffer,
        uint nNumberOfCharsToRead,
        out uint lpNumberOfCharsRead,
        IntPtr pInputControl);
}

sealed class ProductSummary
{
    public string Query { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Price { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
}

sealed class ProductDetails
{
    public string Query { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Price { get; set; } = string.Empty;
    public string ProductUrl { get; set; } = string.Empty;
    public string ProductPageTitle { get; set; } = string.Empty;
    public string MarketArticle { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string SellerName { get; set; } = string.Empty;
    public string SellerUrl { get; set; } = string.Empty;
    public string SellerLegalName { get; set; } = string.Empty;
    public string SellerInn { get; set; } = string.Empty;
    public string SellerOgrn { get; set; } = string.Empty;
    public string ErrorMessage { get; set; } = string.Empty;
    public Dictionary<string, string> Specs { get; } = new(StringComparer.OrdinalIgnoreCase);

    public static ProductDetails FromFailedProduct(ProductSummary product, string errorMessage)
    {
        return new ProductDetails
        {
            Query = product.Query,
            Title = product.Title,
            Price = product.Price,
            ProductUrl = product.Url,
            ErrorMessage = string.IsNullOrWhiteSpace(errorMessage)
                ? string.Empty
                : Regex.Replace(errorMessage.Trim(), @"\s+", " ")
        };
    }
}

sealed class SellerInfo
{
    public string Name { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
}

sealed class SellerLegalInfo
{
    public string LegalName { get; set; } = string.Empty;
    public string Inn { get; set; } = string.Empty;
    public string Ogrn { get; set; } = string.Empty;

    public bool HasAnyValue =>
        !string.IsNullOrWhiteSpace(LegalName) ||
        !string.IsNullOrWhiteSpace(Inn) ||
        !string.IsNullOrWhiteSpace(Ogrn);
}
