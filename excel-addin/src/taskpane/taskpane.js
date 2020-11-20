/* global console, document, Excel, Office */

const AUTH_KEY = 'Bearer ***KEY HERE***';
const ORGANIZATION = '***ORG HERE***'

Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log('Sorry, this add-in uses Excel.js APIs that are not available in your version of Office.');
        }

        // Assign event handlers and other initialization logic.
        document.getElementById("generate-random-topic").onclick = generateRandomTopic;
        document.getElementById("suggest-header").onclick = suggestHeader;
        document.getElementById("complete-table").onclick = completeTable;
        document.getElementById("toggle-extra-info-hidden").onclick = toggleExtraInfo;
        document.getElementById("toggle-extra-info-visible").onclick = toggleExtraInfo;

        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
    }
    localStorage.clear();
});

import sha224 from 'js-sha256';
var request = require('sync-request');


var CELL_SEPARATOR = ' | '
var CONTEXT = `Please build a table summarizing wars in the 20th century
| Name | Start | End |
| Gulf War | 1991 | 2003 |
| Vietnam War | 1955 | 1975 |
| Korean War | 1950 | 1953|
| World War II | 1941 | 1945 |
| Spanish Civil War | 1936 | 1939 |
| World War I | 1914 | 1918 |
| Russo-Japanese War | 1904 | 1905 |
| Boer War | 1899 | 1902 |
| Boxer Rebellion | 1899 | 1901 |
| Mexican Revolution | 1910 | 1920 |
| Russo-Turkish War | 1877 | 1878 |
| Franco-Prussian War | 1870 | 1871 |
| Franco-Austrian War | 1859 | 1866 |

Please build a table summarizing top websites
| Site | Alexa Rank | Revenue |
| Google | 1 | $66.89 billion |
| Facebook | 2 | $15.09 billion |
| Youtube | 3 | $5.87 billion |
| Yahoo! | 4 | $3.81 billion |
| Wikipedia | 5 | $1.88 billion |
| Amazon | 6 | $1.88 billion |
| eBay | 7 | $1.87 billion |
`

function _hash(...args) {
    return sha224(args.join(""));
}


function _strip(string) {
    return string.replace(/(^[ '\^\$\*#&]+)|([ '\^\$\*#&]+$)/g, '');
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function setWithExpiry(key, value, ttls) {
    const now = new Date()

    const item = {
	value: value,
	expiry: now.getTime() + ttls * 1000 // convert s to ms
    }
    localStorage.setItem(key, JSON.stringify(item))
}

function getWithExpiry(key) {
    const itemStr = localStorage.getItem(key)
    // if the item doesn't exist, return null
    if (!itemStr) {
	return null
    }
    const item = JSON.parse(itemStr)
    const now = new Date()
    // compare the expiry time of the item with the current time
    if (now.getTime() > item.expiry) {
	// If the item is expired, delete the item from storage
	// and return null
	localStorage.removeItem(key)
	return null
    }
    return item.value
}


function _build_field_query_context_string({topic, fields=[]}) {
    var field_string;
    if (!fields.length) {
        field_string = '|';
    }
    else {
        field_string = `| ${fields.join(CELL_SEPARATOR)} |`;
    }
    const TOPIC_QUERY = `\nPlease build a table summarizing ${topic}\n`;
    return `${CONTEXT}${TOPIC_QUERY}${field_string}`;
}

function _build_completion_query_context_string({topic, fields, completions=[], partial=[], side_info=[]}) {
    // TODO assert for malformed completions (incomplete tables)
    const field_string = `| ${fields.join(CELL_SEPARATOR)} |\n`;
    var completion_string = '';
    var side_info_string = '';
    if (!completions.length) {
        completion_string = '|'
    }
    else {
        completion_string = completions.map(x => `| ${x.join(CELL_SEPARATOR)} |`).join("\n") + "\n|";
    }

    if (partial.length) {
        completion_string += ` ${partial.join(CELL_SEPARATOR)}`;
    }

    if (side_info.length) {
        side_info_string = '\n' + side_info.join('\n') + '\n\n';
    }
    else {
        side_info_string = '';
    }
    const TOPIC_QUERY = `Please build a table summarizing ${topic}\n`;
    return {context_string: `${CONTEXT}${side_info_string}${TOPIC_QUERY}${field_string}`, existing_completions: `${completion_string}`};
}

function _do_fetch_json(method, url, options) {
    const res = request(method, url, options);
    return JSON.parse(res.getBody());
}

async function cached_request({context_string, logprobs, length, temperature=0.0}) {
    if (context_string === undefined) {
        context_string = _build_field_query_context_string({topic: 'Edible fruits'});
    }
    if (logprobs === undefined) {
        logprobs = 0;
    }
    if (length === undefined) {
        length = 24;
    }
    const request_hash = _hash(context_string, logprobs, length);
    var cached_result = getWithExpiry(request_hash);
    if (cached_result === null) {
        // TODO handle non-200 code
        cached_result = make_request({context_string: context_string, logprobs: logprobs, length: length, temperature: temperature});
    }
    else {
        await sleep(Math.floor(Math.random() * 500) + 250);
        function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
    }
    try {
        setWithExpiry(request_hash, cached_result, 1800);
    }
    catch(err) {
        console.log('Issue caching, possible too large: ', err, cached_result.length);
        console.log(JSON.stringify(cached_result).length);
    }
    return cached_result;
}

function make_request({context_string, logprobs, length, temperature=0.0}) {
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': AUTH_KEY,
        'OpenAI-Organization': ORGANIZATION,
    };
    var payload = {
        "completions": 1,
        "context": context_string,
        "length": length,
        "logprobs": logprobs,
        "stream": false,
        "temperature": temperature,
        "top_p": 1,
    };
    // Practical example
    var response = _do_fetch_json('post', 'https://api.openai.com/v1/engines/davinci/generate', {
        headers: headers,
        json: payload,
        gzip: false,
    });
    return response;
}

// given a topic, and a (possibly empty, possibly partial) list of fields,
// produce a new list of fields + possible extra completions
async function suggest_fields({topic, fields=[]}) {
    const context_string = _build_field_query_context_string({topic: topic, fields: fields});
    const response_json = await cached_request({context_string: context_string, logprobs: 5, length: 48});
    const text_offsets = response_json['data'][0]['text_offset'];
    const texts = response_json['data'][0]['text'];
    const logits = response_json['data'][0]['top_logprobs'];

    // first index after context
    var context_idx = text_offsets.indexOf(context_string.length);
    // newline at end of sampled header
    var endheader_idx = texts.indexOf('\n', context_idx);
    // pull out the actual fields, the -1 is for the ending |
    var sampled_fields = _strip(texts.slice(context_idx, endheader_idx-1).join("")).split(CELL_SEPARATOR);

    // # we need to join onto end tokens
    if (fields.length && _strip(texts[context_idx-1]) == fields[fields.length-1]) {
        sampled_fields = [fields[fields.length-1] + sampled_fields[0], ...sampled_fields.slice(1)];
        fields = fields.slice(0, -1);
    }

    // pull the other possible options for the endheader newline. these should be
    // other possible fields, in likelihood order (though it will only be the
    // first token)
    const sorted_fields = Object.entries(logits[endheader_idx][0]).sort((a, b) => { return b[1] - a[1] });
    const extra_fields = sorted_fields.filter(x => x[0] != '\n' && x[0] != ' ').map(x => [_strip(x[0]), x[1]]);

    return {fields: [...fields, ...sampled_fields].filter(x => x.length), extra_fields: extra_fields};
    //return {fields: [...fields, ...sampled_fields].filter(x => x.length), extra_fields: []};
}


function _advance_line(start_idx, texts) {
    const endline_idx = texts.indexOf('\n', start_idx);
    if (endline_idx == -1) {
        return {result: "", next_idx: -1};
    }
    return {result: _strip(texts.slice(start_idx, endline_idx-1).join("")).split(CELL_SEPARATOR), next_idx: endline_idx};
}


function _merge_field_suggestions(fields_obj, num) {
    // produce a specific number of field suggestions
    const fields = fields_obj['fields'];
    const extra = fields_obj['extra_fields'];
    var merged_list = [...fields,...extra.filter(x => x[0].length > 2).map(x => x[0])];
    return merged_list.slice(0, num);
}

function _maybe_prune_completions(completions_obj, fields, num) {
    // produce a specific number of completion suggestions
    var completions = completions_obj['result'];
    var partial = completions_obj['partial'];

    for(let i = 0; i < completions.length; i++) {
        if (completions[i].length < fields) {
            return {result: completions.slice(0, i), partial: completions[i]};
        }
        if (i + 2 >= num) {  // adding 1 for header row
            return {result: completions.slice(0, num - 1), partial: []};
        }
    }

    return completions_obj;
}

// given a topic, a list of fields, and a (possibly empty,
// possibly partial) list of completions, produces more completions. Passing
// partial completions avoids repeats
async function suggest_completions({topic, fields, completions=[], partial=[], side_info=[], temperature=0.0}) {
    const query_obj = _build_completion_query_context_string({topic: topic, fields: fields, completions: completions, partial: partial, side_info: side_info});
    const context_string = query_obj["context_string"];
    const existing_completions = query_obj["existing_completions"];
    const response_json = await cached_request({context_string: context_string + existing_completions, logprobs: 0, length: 100, temperature: temperature});

    // TODO could use logits here to produce hybrid answers
    const text_offsets = response_json['data'][0]['text_offset'];
    const texts = response_json['data'][0]['text'];

    // first index after context
    var start_idx = text_offsets.indexOf(context_string.length + 1);  // always ends on a |
    var parsed_results = new Array();
    while (true) {
        var next_line_obj = _advance_line(start_idx, texts)
        var result = next_line_obj['result'];
        var next_idx = next_line_obj['next_idx'];
        if (next_idx == -1) {
            break;
        }
        // hack: sometimes we get empty completions :(
        if (result.length && result != ['']) {
            parsed_results.push(result);
        }
        start_idx = next_idx + 2  // skipping newline and starting cell separator
    }
    // parse the last partial line
    var last_line_tokens = texts.slice(start_idx);
    if (last_line_tokens.length && last_line_tokens[last_line_tokens.length - 1] == ' |') {
        last_line_tokens = last_line_tokens.slice(0, last_line_tokens.length - 1);
    }
    var last_partial = _strip(last_line_tokens.join("")).split(CELL_SEPARATOR);
    return {result: parsed_results, partial: last_partial.filter(x => !!(x))};
}


function errorHandlerFunction(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    throw error;
}


function generateRandomTopic() {
    // TODO select random element from list of topics, fill in to text field
    const items = ['Edible fruits', 'US Presidents', 'Unicorn startups'];
    const item = items[Math.floor(Math.random() * items.length)];
    //const item = items[1];
    var topic_box = document.getElementById("topic-box");
    topic_box.value = item;
}

function try_fetch_topic() {
    var topic = document.getElementById("topic-box").value;
    if (topic === '') {
        return null;
    }
    return topic;
}

function try_fetch_side_info() {
    var topic = document.getElementById("side-info-box").value;
    if (topic === '') {
        return null;
    }
    return topic;
}

function suggestHeader() {
    var topic = try_fetch_topic();
    if (topic === null) {
        showToast("Please specify a topic (or click \"Generate a Random Topic\")")
        return;
    }
    showToast(`Suggesting headers for ${topic}`);
    Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = context.workbook.getSelectedRange();
        range.load("values,rowIndex,columnIndex,rowCount,columnCount");
        await context.sync();

        range.numberFormat = "@";
        const boldProps = {
            format: {
                font: {
                    bold: true
                }
            }
        };
        range.set(boldProps);

        // TODO check length == 1
        const fields = range.values;
        var fields_obj = await suggest_fields({topic: topic, fields: fields[0].filter(x => String(x).length)});
        var result_fields = _merge_field_suggestions(fields_obj, fields[0].length);
        sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, 1, result_fields.length).values = [result_fields];
        return context.sync();
    }).catch(errorHandlerFunction);
}

function completeTable() {
    var topic = try_fetch_topic();
    if (topic === null) {
        showToast("Please specify a topic (or click \"Generate a Random Topic\")")
        return;
    }
    showToast(`Completing a table for ${topic}...`, 4000);
    Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = context.workbook.getSelectedRange();
        range.load("values,rowIndex,columnIndex,rowCount,columnCount");
        await context.sync();

        if (range.rowCount < 2) {
            showToast("Please select a table area with at least 2 rows before clicking \"Complete Table\"");
            return;
        }

        const fields = range.values[0].filter(x => String(x).length);
        if (!fields.length) {
            showToast("Couldn't find header row - try clicking \"Suggest Header\" first.");
            return;
        }

        var completions = range.values.slice(1, range.rowCount);
        var partial = [];
        for(let i = 0; i < completions.length; i++) {
            var possible_partial = completions[i].filter(x => String(x).length);
            if (possible_partial.length < fields.length) {
                completions = completions.slice(0, i);
                partial = possible_partial;
                break;
            }
        }
        var side_info = try_fetch_side_info();
        var completions_obj = null;
        var temperature = 0.0;
        if (side_info === null) {
            completions_obj = await suggest_completions({topic: topic, fields: fields, completions: completions, partial: partial, temperature: temperature});
        }
        else {
            completions_obj = await suggest_completions({topic: topic, fields: fields, completions: completions, partial: partial, side_info: [side_info], temperature: temperature});
        }

        // set area as unbolded text
        range.numberFormat = "@";
        const unboldProps = {
            format: {
                font: {
                    bold: false
                }
            }
        };
        range.set(unboldProps);
        const boldProps = {
            format: {
                font: {
                    bold: true
                }
            }
        };
        sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, 1, range.columnCount).set(boldProps);

        var max_tries = 3;
        for(let i = 0; i < max_tries; i++) {
            var pruned_completions = _maybe_prune_completions(completions_obj, fields.length, range.rowCount);
            var tmp_results = pruned_completions["result"].slice();
            tmp_results.unshift(fields);
            sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, tmp_results.length, range.columnCount).values = tmp_results;
            await context.sync();
            if (pruned_completions['partial'].length && pruned_completions['partial'][0].startsWith('Please build a table')) {
                break;
            }
            if (tmp_results.length == range.rowCount) {
                break;
            }
            if (i < max_tries - 1) {
                showToast(`Completing a table for ${topic}...`);
                temperature += 0.7;
                completions_obj = await suggest_completions({topic: topic, fields: fields, completions: pruned_completions['result'], partial: pruned_completions['partial'], temperature: temperature});
            }
        }
        if (pruned_completions["result"].length + 1 != range.rowCount) {
            showToast("Could not complete table in one go; this usually means the selected area is too big. You can try clicking \"Complete Table\" again.")
        }
        return context.sync();
    }).catch(errorHandlerFunction);
}


function showToast(msg, time_s=2000) {
    var elem = document.getElementById("snackbar");
    elem.className = "show";
    elem.innerText = msg;
    setTimeout(function(){ elem.className = elem.className.replace("show", ""); }, time_s);
}


var extra_info_visible = false;

function toggleExtraInfo() {
    var elem = document.getElementById("side-info-container");
    if (extra_info_visible) {
        elem.className = "side-info-container-hidden topic-form-element";
        document.getElementById("toggle-extra-info-visible").style.display = "none";
        document.getElementById("toggle-extra-info-hidden").style.display = "inline-block";
        extra_info_visible = false;
    }
    else {
        elem.className = "side-info-container-visible topic-form-element";
        document.getElementById("toggle-extra-info-visible").style.display = "inline-block";
        document.getElementById("toggle-extra-info-hidden").style.display = "none";
        extra_info_visible = true;
    }
}
