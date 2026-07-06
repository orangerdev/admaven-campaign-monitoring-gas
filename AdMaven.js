class AdMaven {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.baseUrl = 'https://advertisers.ad-maven.com/api/public';
  }

  /**
   * Fetch report data from AdMaven API
   * @param {string} fromDate - Format: YYYY-MM-DD
   * @param {string} toDate - Format: YYYY-MM-DD
   * @param {string[]} groupBy - Valid values: country_code, device_type, report_date, campaign_id,
   *                             browser_name, operating_system, traffic_type, source_id, sub_source_id
   * @param {string[]} columns - Valid values: redirects, conversions, cost, revenue,
   *                             daily_budget, global_budget, destination_url, campaign_id, name
   * @param {Object} [filters] - Optional filter fields: device_type, campaign_id, browser_name,
   *                             operating_system, country_code, traffic_type, source_id, sub_source_id
   * @returns {Object[]} Array of report row objects
   */
  getReport(fromDate, toDate, groupBy, columns, filters) {
    const url = `${this.baseUrl}/reports`;

    const payload = {
      filters: Object.assign(
        { from_date: fromDate, to_date: toDate },
        filters || {}
      ),
      group_by: groupBy,
      columns:  columns,
    };

    const response = UrlFetchApp.fetch(url, {
      method:      'POST',
      contentType: 'application/json',
      headers: {
        'accept':        'application/json',
        'Authorization': `Bearer ${this.apiKey}`,
      },
      payload:           JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      throw new Error(`AdMaven API error ${statusCode}: ${response.getContentText()}`);
    }

    const json = JSON.parse(response.getContentText());

    // Response structure: { type, message: { columns, rows } }
    if (json.message && Array.isArray(json.message.rows)) {
      return json.message.rows;
    }

    // Fallback for plain array responses
    return Array.isArray(json) ? json : [];
  }

  /**
   * Fetch all campaigns (no ID filter)
   * @param {string} fields - Comma-separated fields to return (e.g. 'name,id')
   * @returns {Object[]} Array of campaign objects
   */
  getAllCampaigns(fields) {
    let url = `${this.baseUrl}/campaign`;
    if (fields) url += `?fields=${fields}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'accept':        'application/json',
        'Authorization': `Bearer ${this.apiKey}`,
      },
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      throw new Error(`AdMaven API error ${statusCode}: ${response.getContentText()}`);
    }

    const json = JSON.parse(response.getContentText());
    return Array.isArray(json) ? json : (json.data || json.message || []);
  }

  /**
   * Fetch option values for a given option type
   * @param {string} type - e.g. 'device_type', 'browser', 'operating_system', 'traffic_type'
   * @returns {Object[]} Array of option objects (e.g. [{ id, name }])
   */
  getOptions(type) {
    const url = `${this.baseUrl}/options/${type}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'accept':        'application/json',
        'Authorization': `Bearer ${this.apiKey}`,
      },
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      throw new Error(`AdMaven API error ${statusCode}: ${response.getContentText()}`);
    }

    const json = JSON.parse(response.getContentText());
    return Array.isArray(json) ? json : (json.data || json.message || []);
  }

  /**
   * Fetch campaigns by IDs
   * @param {number[]} ids - Array of campaign IDs
   * @param {string} fields - Comma-separated fields to return (e.g. 'name,enable,rates')
   * @returns {Object[]} Array of campaign objects
   */
  getCampaigns(ids, fields) {
    if (!ids || ids.length === 0) return [];

    let url = `${this.baseUrl}/campaign?id=${ids.join(',')}`;
    if (fields) url += `&fields=${fields}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'accept':        'application/json',
        'Authorization': `Bearer ${this.apiKey}`,
      },
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      throw new Error(`AdMaven API error ${statusCode}: ${response.getContentText()}`);
    }

    const json = JSON.parse(response.getContentText());
    return Array.isArray(json) ? json : (json.data || json.message || []);
  }
}
