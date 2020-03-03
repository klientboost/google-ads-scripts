/**
 * Converts "exact (close variant)" search queries into negative keywords
 *
 * @author spencer@klientboost.com
 * @version 1.0
 *
 * -------------------------------- INSTRUCTIONS --------------------------------
 *
 * 1. Download this script and upload it to your Google Ads account under
 *    Tools & Settings > Bulk Actions > Scripts
 *
 *    More info: https://klientboost.com/ppc/adwords-scripts/
 *
 *
 * 2. Review lines 58-61 and edit variables for your settings.
 *
 *        Line 58 - dateRange - Define date range you would like this script to lookback
 *                  for close variant negatives. If you would like this to be an everyday,
 *                  ongoing script, place the date range for years in advance and set
 *                  the script to run daily (we suggest this :).
 *
 *        Line 59 - campaignNameFilter - This is where you will define what campaigns you
 *                  would like this script to work on. Our methodology is to have single
 *                  keyword ad groups within exact match and broad modified campaigns. So
 *                  since this script adds search terms as negatives on the ad group level,
 *                  we can know that no traffic is being cut off from any other keywords
 *                  or campaigns when this script is applied to EXACT MATCH CAMPAIGNS
 *                  ONLY. So if you would like to use this script to full effect, we
 *                  suggest a naming convention that defines only exact match campaigns.
 *
 *        Line 60 - email - Enter the email you would like the report to send to. It will
 *                  contain a list of all "exact (close variant)" search queries and the ad
 *                  groups where they were found.  It will also show count of conversions if
 *                  any there were any in the time period.
 *
 *        Line 61 - emailOnly - Determines whether you would like these negatives to be added
 *                  automatically or only sent to your email!
 *
 *                  emailOnly = true;   # Only an email summary is sent, no changes to your
 *                                      # account.
 *
 *                  emailOnly = false;  # All "exact (close variant)" will be added as exact match
 *                                      # negative keywords to your account automatically. An email
 *                                      # summary will also be sent.
 *
 *
 * 3. Define how often you would like the script to run. We have this script
 *    running daily at 2pm because we want this script to be alive and active
 *    every day to keep our exact match keywords… exact match!
 *
 *    This setting is found in the "Frequency" column when viewing Google Ads scripts.
 *
 */
function main() {
  /**
   * 0. Set custom rules like date range and campaign name filter
   */
  var dateRange = "LAST_7_DAYS"; // E.g. "LAST_7_DAYS" | "20190101,20200228"
  var campaignNameFilter = ""; // E.g. "*KB" (leave blank to process all campaigns)
  var email = "spencer@klientboost.com"; // E.g. "spencer@klientboost.com"
  var emailOnly = false; // true | false

  /**
   * 1. Get a list of exact match keywords
   */
  var exactMatchKeywords = getExactMatchAdGroupKeywords(dateRange);

  /**
   * 2. Get a list of close variant search queries by ad group
   */
  var closeVariantSearchQueries = getExactCloseVariantSearchQueries(
    dateRange,
    campaignNameFilter
  );

  /**
   * 3. Iterate through ad groups and add negative keywords to them
   */
  var log = addNegativeKeywordsToAdGroups(
    closeVariantSearchQueries,
    exactMatchKeywords,
    emailOnly
  );

  /**
   * 4. Email the log to the email address provided
   */
  emailLog(log, email, emailOnly, dateRange);
}

/**
 * Ads search queries as negative keywords and emails the results
 * @param {object} closeVariantSearchQueries object containing search queries, keyed by adGroupId
 * @param {Array<string>} exactMatchKeywords an array exact match keyword keys to validate against
 * @param {bool} emailOnly as string filter for the CampaignName field
 * @returns {Array<string>} an array of log strings representing updates
 */
function addNegativeKeywordsToAdGroups(
  closeVariantSearchQueries,
  exactMatchKeywords,
  emailOnly
) {
  var log = [];
  var adGroupIds = Object.keys(closeVariantSearchQueries);
  var adGroupIterator = AdsApp.adGroups()
    .withIds(adGroupIds)
    .get();
  while (adGroupIterator.hasNext()) {
    var adGroup = adGroupIterator.next();
    var adGroupId = adGroup.getId();

    if (closeVariantSearchQueries[adGroupId]) {
      var adGroupName = closeVariantSearchQueries[adGroupId].adGroup;
      var campaignName = closeVariantSearchQueries[adGroupId].campaign;
      var searchQueries = closeVariantSearchQueries[adGroupId].searchQueries;
      log.push("");
      log.push("Campaign: " + campaignName);
      log.push("Ad Group: " + adGroupName);
      log.push("Search Queries: " + adGroupName);

      searchQueries.forEach(function(searchQuery) {
        var key = searchQuery.key;
        var query = searchQuery.searchQuery;
        var conversions = searchQuery.conversions;
        // Only create negative keywords with corrosponding exact match keywords
        if (exactMatchKeywords.indexOf(key) === -1) {
          return;
        }
        if (!emailOnly) {
          adGroup.createNegativeKeyword("[" + query + "]");
        }
        var line = " - " + query;
        if (conversions >= 1) {
          line =
            line +
            " (" +
            conversions +
            " conversion" +
            (conversions !== 1 ? "s" : "") +
            ")";
        }
        log.push(line);
      });
    }
  }
  return log;
}

/**
 * Get report data from the Google Ads KEYWORDS_PERFORMANCE_REPORT
 * @param {string} dateRange a Google Ads QL Date Range
 * @returns {object} a Google Ads report result
 */
function getKeywordPerformanceReport(dateRange) {
  // prettier-ignore
  return (AdsApp.report(
    "SELECT Id, Criteria, AdGroupId, AdGroupName \
      FROM KEYWORDS_PERFORMANCE_REPORT \
      WHERE Impressions > 0 AND KeywordMatchType = EXACT \
      DURING " + dateRange + " \
      "
  )).rows();
}

/**
 * Get report data from the Google Ads SEARCH_QUERY_PERFORMANCE_REPORT
 * @param {string} dateRange a Google Ads QL DateRange
 * @param {string} campaignNameFilter as string filter for the CampaignName field
 * @returns {object} a Google Ads report result
 */
function getSearchQueryPerformanceReport(dateRange, campaignNameFilter) {
  // prettier-ignore
  return (AdWordsApp.report(
    "SELECT Query, AdGroupId, AdGroupName, CampaignId, CampaignName, KeywordId, \
      KeywordTextMatchingQuery, Impressions, Clicks, Conversions, QueryMatchTypeWithVariant \
      FROM SEARCH_QUERY_PERFORMANCE_REPORT \
      WHERE CampaignName CONTAINS_IGNORE_CASE '" + campaignNameFilter + "' \
      DURING " + dateRange + " \
      "
  )).rows();
}

/**
 * Get an array of unique adGroup::keywordId keys for exact match keywords
 * @param {string} dateRange a Google Ads QL Date Range
 * @returns {Array<string>} an array of unique adGroup::keywordId keys
 */
function getExactMatchAdGroupKeywords(dateRange) {
  var keywords = [];
  var keywordPerformanceReport = getKeywordPerformanceReport(dateRange);
  while (keywordPerformanceReport.hasNext()) {
    var keyword = keywordPerformanceReport.next();
    var Id = keyword.Id;
    var AdGroupId = keyword.AdGroupId;
    var key = AdGroupId + "::" + Id;

    keywords.push(key);
  }
  return keywords;
}

/**
 *
 * @param {string} dateRange a Google Ads QL Date Range
 * @param {string} campaignNameFilter as string filter for the CampaignName field
 * @returns {Array<object>} an array of objects containing adGroup, campaign, and searchQueries with performance data
 */
function getExactCloseVariantSearchQueries(dateRange, campaignNameFilter) {
  var adGroupSearchQueries = {};
  var searchQueryPerformanceReport = getSearchQueryPerformanceReport(
    dateRange,
    campaignNameFilter
  );
  while (searchQueryPerformanceReport.hasNext()) {
    var searchQuery = searchQueryPerformanceReport.next();
    var AdGroupId = parseInt(searchQuery.AdGroupId);
    var AdGroupName = searchQuery.AdGroupName;
    var CampaignId = parseInt(searchQuery.CampaignId);
    var CampaignName = searchQuery.CampaignName;
    var KeywordId = parseInt(searchQuery.KeywordId);
    var Impressions = parseInt(searchQuery.Impressions);
    var Clicks = parseInt(searchQuery.Clicks);
    var Conversions = parseInt(searchQuery.Conversions);
    var Query = searchQuery.Query.toLowerCase();
    var KeywordTextMatchingQuery = searchQuery[
      "KeywordTextMatchingQuery"
    ].toLowerCase();
    var QueryMatchTypeWithVariant = searchQuery[
      "QueryMatchTypeWithVariant"
    ].toLowerCase();

    // Test to make sure the search query is not the keyword
    if (
      KeywordTextMatchingQuery !== Query &&
      QueryMatchTypeWithVariant.trim() == "exact (close variant)"
    ) {
      var key = AdGroupId + "::" + KeywordId;

      // First search query for this ad group
      if (!adGroupSearchQueries[AdGroupId]) {
        adGroupSearchQueries[AdGroupId] = {
          searchQueries: [],
          adGroup: AdGroupName,
          campaign: CampaignName
        };
      }

      // Append the search query information and performance data to the ad group array
      adGroupSearchQueries[AdGroupId].searchQueries.push({
        key: key,
        keyword: KeywordTextMatchingQuery,
        searchQuery: Query,
        matchType: QueryMatchTypeWithVariant,
        impressions: Impressions,
        clicks: Clicks,
        conversions: Conversions
      });
    }
  }
  return adGroupSearchQueries;
}

/**
 * Parses the log array and sends an email
 * @param {Array<string>} log an array of log messages
 * @param {string} email an email to send the log to
 * @param {boolean} emailOnly true if the changes have not been implemented by the script
 * @param {string} dateRange the date range used to determine changes
 * @returns {undefined} no return value
 */
function emailLog(log, email, emailOnly, dateRange) {
  var account = AdsApp.currentAccount();
  /**
   * Add header information to the log
   */
  log.unshift(
    "Your ad groups generated these `exact (close variant)` search queries in the `" +
      dateRange +
      "` date range. They have" +
      (emailOnly ? " not" : "") +
      " been added as exact match negative keywords."
  );
  log.unshift("");

  /**
   * Print the log array to the console
   */
  log.forEach(function(line) {
    Logger.log(line);
  });

  /**
   * Email the log to the email address provided
   */
  var subject = [
    "Close Match Script",
    account.getName(),
    account.getCustomerId()
  ].join(" | ");
  var mailResult = MailApp.sendEmail(email, subject, log.join("\n"));
}
