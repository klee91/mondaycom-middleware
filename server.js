/**
 * Monday.com Email Template Generator — Full Stack Server
 * Serves frontend UI + API + Monday webhook listener
 *
 * Required env vars:
 *   MONDAY_API_TOKEN   — Monday.com Personal API Token
 *   ANTHROPIC_API_KEY  — Anthropic API Key
 *   BOARD_ID           — Monday.com Board ID
 *   MONDAY_SIGNING_SECRET — From Monday app settings (for webhook verification)
 */

const express   = require("express");
const cors      = require("cors");
const fetch     = require("node-fetch").default || require("node-fetch");
const FormData  = require("form-data");
const Anthropic = require("@anthropic-ai/sdk").default;
const crypto    = require("crypto");
const path      = require("path");

const app = express();
app.use(cors({ origin: "*", methods: ["GET", "POST", "OPTIONS"], allowedHeaders: ["Content-Type", "Authorization"] }));
app.options("*", cors());
app.use(express.json({ limit: "10mb" }));

const MONDAY_API_URL  = "https://api.monday.com/v2";
const MONDAY_FILE_URL = "https://api.monday.com/v2/file";
const BOARD_ID        = process.env.BOARD_ID;
const anthropic       = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ─────────────────────────────────────────────
// Templates
// ─────────────────────────────────────────────
const TEMPLATES = {
  "AICPA Town Hall Newsletter": `<!doctype html><html><head><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{{Subject}}</title><style type="text/css">@import url('https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap');table{border-collapse:collapse}table td{border-collapse:collapse;font-weight:300!important}table td strong{font-weight:700!important}body{background-color:#fff;margin:0}.appleLinks a{color:inherit!important;text-decoration:underline}a{text-decoration:underline;color:#86387f}</style></head><body style="margin:0px;padding:0px;" yahoo="fix" bgcolor="#ffffff"><table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" style="table-layout:fixed;" width="100%"><tbody><tr><td align="center" bgcolor="#ffffff" style="margin:0 auto!important;min-width:680px;padding:0;width:680px;" valign="top" width="680"><div align="center"><table align="center" border="0" cellpadding="0" cellspacing="0" style="width:680px;max-width:680px;" width="680"><tbody><tr><td align="center" valign="top"><table cellpadding="0" cellspacing="0" style="max-width:680px;" width="680"><tbody><tr><td><table align="center" border="0" cellpadding="0" cellspacing="0" style="max-width:680px" width="680"><tbody><tr><td align="center" style="padding:10px 25px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;color:#231f20;font-weight:300;text-align:center;" valign="top"><p style="font-family:'Roboto',Arial,Helvetica,sans-serif;color:#000000;font-size:16px;font-weight:300;"><strong style="font-weight:700;"><a href="https://www.aicpa-cima.com/cpe-learning/webcast/aicpa-town-hall-series" style="text-decoration:underline;color:#86387f;" target="_blank">{{PreheaderText}}</a></strong></p></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table cellpadding="0" cellspacing="0"><tbody><tr><td><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation"><tbody><tr><td align="center"><a href="https://www.cpa.com/town-hall" target="_blank"><img alt="AICPA Town Hall Series Newsletter" border="0" height="170" src="https://marketing.cpa.com/l/701003/2025-05-14/46x5rg/701003/17472392906lBfnhW5/AICPA_TH_Podcast_Banner.png" style="width:680px;height:170px;" width="680"></a></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table cellpadding="0" cellspacing="0" style="background-color:#FFFFFF;min-width:100%;border-bottom:solid 1px #ccc;" width="100%"><tbody><tr><td style="padding:0px;"><table align="center" bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="680"><tbody><tr><td style="padding:25px 50px 15px 50px;line-height:23px;color:#000;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;font-weight:200;"><p style="font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;font-weight:300;">Hi {{Recipient.FirstName}},</p><p style="font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;font-weight:300;">BODY_CONTENT_HERE</p></td></tr></tbody></table></td></tr></tbody></table></td></tr><!-- START FOOTER --><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" bgcolor="#000000" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;background:#000000;" width="680"><tbody><tr><td align="center" valign="middle"><div style="display:inline-block;margin:0;max-width:50%;min-width:240px;vertical-align:middle;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td style="font-family:Arial,Verdana;font-size:20px;line-height:24px;color:#72246c;padding:10px 50px;text-align:left;"><a href="https://www.cpa.com" style="color:#ffffff;text-decoration:none;" target="_blank">CPA.com</a></td></tr></tbody></table></div></td></tr></tbody></table><table align="center" bgcolor="#eeeeee" cellpadding="0" cellspacing="0" style="width:680px;padding:0 25px;" width="680"><tbody><tr><td style="text-align:center;vertical-align:top;font-size:0;padding:30px 0px 0px"><div><div style="display:inline-block;vertical-align:top;max-width:330px;"><table align="right" border="0" cellpadding="0" cellspacing="0" height="180" width="330"><tbody><tr><td align="right" style="padding:0px 20px 0px 50px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:12px;color:#63656b;line-height:100%;font-weight:300;text-align:left;" valign="top"><span>1345 Avenue of the Americas, 27th Floor</span><br><span>New York, NY 10019</span><br><span>888.777.7077</span></td></tr></tbody></table></div></div></td></tr></tbody></table></td></tr><!-- END FOOTER --></tbody></table></div></td></tr></tbody></table></body></html>`,

  "DOTCPA General": `<!doctype html><html><head><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{{Subject}}</title><style type="text/css">@import url('https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap');body{background-color:#ffffff;margin:0}table{border-collapse:collapse}</style></head><body bgcolor="#ffffff" style="margin:0px;padding:0px;" yahoo="fix"><table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" style="table-layout:fixed;" width="100%"><tbody><tr><td align="center" bgcolor="#ffffff" style="margin:0 auto!important;min-width:680px;padding:0;width:680px;" valign="top" width="680"><div align="center"><table align="center" border="0" cellpadding="0" cellspacing="0" style="width:640px;max-width:640px;" width="640"><tbody><tr><td align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" style="max-width:680px;border-bottom:1px solid #cccccc" width="100%"><tbody><tr><td align="center"><a href="https://register.domains.cpa/"><img alt=".CPA - A service of AICPA and CPA.com" border="0" height="120" src="https://marketing.cpa.com/l/701003/2020-09-15/n6drm/701003/77696/dotCPA_service_email_header.png" style="width:680px;height:120px;" width="680"></a></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px" width="100%"><tbody><tr><td style="padding:25px 50px 0px 50px;font-family:'Roboto',Arial,Verdana;font-weight:300;font-size:17px;line-height:23px;color:#63656b;"><p style="font-family:'Roboto',Arial,sans-serif;font-size:16px;color:#000000;">Hi {{Recipient.FirstName}},</p><p style="font-family:'Roboto',Arial,sans-serif;font-size:16px;color:#000000;">BODY_CONTENT_HERE</p><table border="0" cellpadding="0" cellspacing="0" width="100%"><tbody><tr><td align="center"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:separate!important;" width="300"><tbody><tr><td style="background-color:#72246C;color:#FFFFFF;font-size:15px;padding:10px;border-radius:3px;font-family:Arial,Helvetica,sans-serif;text-align:center;"><a href="https://register.domains.cpa/ga/#start" style="color:#FFFFFF;text-decoration:none;">BUTTON_TEXT</a></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr><!-- START FOOTER --><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" bgcolor="#000000" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="100%"><tbody><tr><td align="center" style="font-size:0;" valign="middle"><div style="display:inline-block;margin:0;max-width:70%;min-width:430px;vertical-align:middle;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td style="font-family:'Roboto',Arial,Verdana;font-size:16px;font-weight:300;color:#72246c;padding:5px 0px 5px 30px;text-align:left;"><a href="http://www.cpa.com/dotcpa" style="color:#ffffff;text-decoration:none;">Domains.CPA</a></td></tr></tbody></table></div></td></tr></tbody></table><table align="center" bgcolor="#eeeeee" cellpadding="0" cellspacing="0" style="max-width:680px;" width="680"><tbody><tr><td style="padding:0px 25px 7px 35px;font-family:'Roboto',Arial;font-size:11px;color:#000;line-height:14px;text-align:left;">JOB_NUMBER_HERE</td></tr></tbody></table></td></tr><!-- END FOOTER --></tbody></table></div></td></tr></tbody></table></body></html>`,

  "CPACOM General": `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns:mso="urn:schemas-microsoft-com:office:office" xmlns:msdt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"><head><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{{Subject}}</title><style type="text/css">@import url('https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap');table{border-collapse:collapse}table td{border-collapse:collapse;font-weight:300!important}table td strong{font-weight:700!important}body{background-color:#fff;margin:0}.appleLinks a{color:inherit!important;text-decoration:underline}a{text-decoration:underline;color:#86387f}</style><!--[if mso]><style type="text/css">table{mso-table-lspace:0pt;mso-table-rspace:0pt;}</style><![endif]--><!--[if gte mso 9]><xml><mso:CustomDocumentProperties><mso:_dlc_DocId msdt:dt="string">DP34VC2UU43T-1062633695-1</mso:_dlc_DocId><mso:_dlc_DocIdItemGuid msdt:dt="string">9563557e-3ba0-4891-bdc4-c4c0cef639fb</mso:_dlc_DocIdItemGuid><mso:_dlc_DocIdUrl msdt:dt="string">https://hqcpa.sharepoint.com/sites/NewProductMarketing/_layouts/15/DocIdRedir.aspx?ID=DP34VC2UU43T-1062633695-1, DP34VC2UU43T-1062633695-1</mso:_dlc_DocIdUrl><mso:docLang msdt:dt="string">en</mso:docLang></mso:CustomDocumentProperties></xml><![endif]--></head><body style="margin:0px;padding:0px;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;" yahoo="fix" bgcolor="#ffffff"><table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" style="table-layout:fixed;" width="100%"><tbody><tr><td align="center" bgcolor="#ffffff" pardot-data="" style="background:rgb(255,255,255);margin:0 auto !important;min-width:680px;padding:0;width:680px;" valign="top" width="680"><div align="center"><table align="center" border="0" cellpadding="0" cellspacing="0" class="container" style="width:680px;max-width:680px;" width="680"><tbody><tr><td align="center" valign="top"><table cellpadding="0" cellspacing="0" class="stylingblock-content-wrapper" style="width:680px;padding:0;" width="680"><tbody><tr><td align="center" class="fnt8" style="padding:0px;border-bottom:solid 1px #CCC;" valign="top"><img alt="CPA.com Header" border="0" height="125" src="https://marketing.cpa.com/l/701003/2022-09-19/3j68h5/701003/1663614409hZUzc2TK/cpacom_header.jpg" style="width:680px;height:125px;border-width:0px;border-style:solid;" width="680"></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table cellpadding="0" cellspacing="0" class="stylingblock-content-wrapper" style="background-color:#FFFFFF;min-width:100%;" width="100%"><tbody><tr><td class="stylingblock-content-wrapper camarker-inner" style="padding:0px;"><!-- A2 (callout block area) : BEGIN --><table align="center" bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="680"><tbody><tr><td class="mtext" dir="ltr" style="padding:20px 50px 0px 50px;text-align:left;color:#000;line-height:22px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;"><p style="font-family:Roboto,Arial,sans-serif;font-size:16px;font-weight:300;">Hi {{Recipient.FirstName}},</p><p style="font-family:Roboto,Arial,sans-serif;font-size:16px;font-weight:300;">BODY_CONTENT_HERE</p></td></tr><tr><td class="mtext" dir="ltr" style="padding:20px 50px 0px 50px;text-align:left;color:#000;line-height:22px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:540px;" width="540px"><tbody></tbody></table></td></tr></tbody></table><!-- A2 (callout block area) : END --></td></tr></tbody></table></td></tr><tr><td class="mtext" dir="ltr" style="padding:20px 50px 0px 50px;text-align:left;color:#000;line-height:22px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:540px;" width="540px"><tbody></tbody></table></td></tr><!-- START FOOTER --><tr><td align="center" bgcolor="#ffffff" valign="top"><table cellpadding="0" cellspacing="0" class="stylingblock-content-wrapper" style="min-width:100%;" width="100%"><tbody><tr><td class="stylingblock-content-wrapper camarker-inner"><!-- Black Footer Bar : BEGIN --><!--[if mso]><table role="presentation" border="0" cellspacing="0" cellpadding="0" align="center" width="680" style="width:680px;"><tr><td align="center" valign="middle" width="680" style="width:680px;"><![endif]--><table align="center" bgcolor="#0F206C" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="100%"><tbody><tr><td align="center" valign="middle"><!--[if mso]><table role="presentation" border="0" cellspacing="0" cellpadding="0" align="center" width="680" style="width:680px;"><tr><td align="left" valign="middle" width="340" style="width:340px;"><![endif]--><div class="stack-column" style="display:inline-block;margin:0;max-width:50%;min-width:240px;vertical-align:middle;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td class="center-on-narrow" dir="ltr" style="font-family:Arial,Verdana;font-size:18px;line-height:24px;color:#72246c;padding:10px 50px 10px 50px;text-align:left;"><a data-linkto="http://" href="https://www.cpa.com" style="color:#ffffff;text-decoration:none;" target="_blank" title="">CPA.com/</a></td></tr></tbody></table></div><!--[if mso]></td><td align="left" valign="middle" width="340" style="width:340px;"><![endif]--><div class="center-on-narrow" style="display:inline-block;margin:0;max-width:50%;min-width:240px;vertical-align:middle;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td align="center" class="m_icons" style="padding:10px 50px 10px 50px;text-align:center;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="140"><tbody><tr><td align="center"><a data-linkto="http://" href="https://www.facebook.com/CPAdotcom/" target="_blank" title=""><img alt="facebook icon" border="0" height="30" src="https://marketing.cpa.com/l/701003/2021-08-11/2cn1xr/701003/1628700105mpymkZuz/fb_icon_80.png" style="width:30px;height:30px;border-width:0px;border-style:solid;" width="30"></a></td><td align="center"><a data-linkto="http://" href="https://twitter.com/cpacom" target="_blank" title=""><img alt="twitter" border="0" height="30" src="https://marketing.cpa.com/l/701003/2021-08-11/2cn1xy/701003/1628700105yFDD4LKA/twitter_icon_80.png" style="display:block;width:30px;height:30px;border-width:0px;border-style:solid;" width="30"></a></td><td align="center"><a data-linkto="https://" href="https://www.linkedin.com/company/cpa-com/" target="_blank" title=""><img alt="linkedin" border="0" height="30" src="https://marketing.cpa.com/l/701003/2021-08-11/2cn1xt/701003/1628700105zOivQQhh/linkedin_icon_80.png" style="display:block;width:30px;height:30px;border-width:0px;border-style:solid;" width="30"></a></td><td align="center"><a data-linkto="http://" href="https://www.cpa.com/blog" target="_blank" title=""><img alt="rss" border="0" height="30" src="https://marketing.cpa.com/l/701003/2021-08-11/2cn1xw/701003/16287001056gbEBdTg/rss_icon_80.png" style="display:block;width:30px;height:30px;border-width:0px;border-style:solid;" width="30"></a></td></tr></tbody></table></td></tr></tbody></table></div><!--[if mso]></td></tr></table><![endif]--></td></tr></tbody></table><!--[if mso]></td></tr></table><![endif]--><!-- Black Footer Bar : END --><table align="center" bgcolor="#ffffff" cellpadding="0" cellspacing="0" class="contenttable container" style="max-width:680px;" width="680"><tbody><tr><td style="text-align:center;vertical-align:top;font-size:0;padding:15px 0px"><div dir="ltr"><!--[if (gte mso 9)|(IE)]><table width="680" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" valign="top" width="340"><![endif]--><div class="container" dir="ltr" style="display:inline-block;vertical-align:top;max-width:340px" width="100%"><table align="center" border="0" cellpadding="0" cellspacing="0" class="container" valign="top" width="330"><tbody><tr><td align="center" style="padding:0px 25px 7px 25px;font-family:'Roboto',Arial;font-size:12px;color:#86387f;line-height:100%;mso-line-height-rule:exactly;font-weight:300;text-align:left;" valign="top"><a data-linkto="http://" href="{{{EmailPreferenceCenter_488}}}" rel="nofollow,noreferrer" style="display:inline;font-weight:400;text-decoration:underline;color:#86387f;" target="_blank" title="Manage your preferences"><span class="fallback-font" style="color:#86387f;text-decoration:underline;font-weight:400;">Manage your preferences</span></a></td></tr><tr><td align="center" style="padding:0px 25px 7px 25px;font-family:'Roboto',Arial;font-size:12px;color:#86387f;line-height:100%;mso-line-height-rule:exactly;font-weight:300;text-align:left;" valign="top"><a data-linkto="http://" href="{{{EmailPreferenceCenter_563}}}" rel="nofollow,noreferrer" style="display:inline;font-weight:400;text-decoration:underline;color:#86387f;" target="_blank" title="Unsubscribe"><span class="fallback-font" style="color:#86387f;text-decoration:underline;font-weight:400;">Unsubscribe</span></a></td></tr><tr><td align="center" style="padding:0px 25px 7px 25px;font-family:'Roboto',Arial;font-size:12px;color:#86387f;line-height:100%;mso-line-height-rule:exactly;font-weight:300;text-align:left;" valign="top"><a data-linkto="http://" href="https://www.cpa.com/privacy-policy" style="color:#86387f;text-decoration:underline;display:inline;font-weight:400;" target="_blank" title="Privacy policy"><span class="fallback-font" style="color:#86387f;text-decoration:underline;font-weight:400;">Privacy policy</span></a></td></tr><tr><td align="center" style="padding:0px 25px 7px 25px;font-family:'Roboto',Arial;font-size:12px;color:#86387f;line-height:100%;mso-line-height-rule:exactly;font-weight:300;text-align:left;" valign="top"><a data-linkto="http://" href="https://www.cpa.com/contact" style="color:#86387f;text-decoration:underline;display:inline;font-weight:400;" target="_blank" title="Contact us"><span class="fallback-font" style="color:#86387f;text-decoration:underline;font-weight:400;">Contact us</span></a></td></tr><tr><td align="center" style="padding:0px 25px 7px 25px;font-family:'Roboto',Arial;font-size:12px;color:#86387f;line-height:100%;mso-line-height-rule:exactly;font-weight:300;text-align:left;" valign="top"><a data-linkto="http://" href="https://marketing.cpa.com/l/701003/2023-03-29/431wkq" rel="nofollow,noreferrer" style="display:inline;font-weight:400;text-decoration:underline;color:#86387f;" target="_blank" title="Resub"><span class="fallback-font" style="color:#86387f;text-decoration:underline;font-weight:400;">Resubscribe</span></a></td></tr></tbody></table></div><!--[if (gte mso 9)|(IE)]></td><![endif]--><!--[if (gte mso 9)|(IE)]><td align="center" valign="top" width="340"><![endif]--><div class="container" dir="ltr" style="display:inline-block;vertical-align:top;width:100%;max-width:330px"><table align="center" border="0" cellpadding="0" cellspacing="0" class="container" width="300"><tbody><tr><td align="left" style="padding:0px 25px 0px 25px;font-family:'Roboto',Arial;font-size:12px;color:#000;line-height:18px;mso-line-height-rule:exactly;font-weight:300;text-align:left;" valign="top"><span class="fallback-font" style="font-family:Roboto,Arial;font-size:12px;">CPA.com</span><br><span class="fallback-font" style="font-family:Roboto,Arial;font-size:12px;">1345 Avenue of the Americas, 27th Floor</span><br><span class="fallback-font" style="font-family:Roboto,Arial;font-size:12px;">New York, NY 10105</span><br><span class="fallback-font" style="font-family:Roboto,Arial;font-size:12px;">888.777.7077</span></td></tr></tbody></table></div><!--[if (gte mso 9)|(IE)]></td></tr></table><![endif]--></div></td></tr></tbody></table><table align="center" bgcolor="#ffffff" cellpadding="0" cellspacing="0" class="contenttable container" style="max-width:680px;" width="680"><tbody><tr><td align="center" rowspan="2" style="padding:0px 25px 7px 35px;font-family:'Roboto',Arial;font-size:11px;color:#000;line-height:14px;mso-line-height-rule:exactly;text-align:left;" valign="top">JOB_NUMBER</td></tr></tbody></table></td></tr></tbody></table></td></tr><!-- END FOOTER --></tbody></table></div></td></tr></tbody></table><div style="display:none;white-space:nowrap;font:15px courier;color:#ffffff;">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -</div></body></html>`
};

const AI_SYSTEM_PROMPT = `You are an email template generator. Populate the HTML template with ticket data.
RULES:
- Return ONLY the populated HTML. No explanation, no markdown, no code fences.
- Keep ALL HTML structure, styles, images, and links completely intact.
- CRITICAL: Any existing {{...}} or {{{...}}} tokens in the template are Pardot merge tags. Leave them exactly as-is. Never modify, replace, or remove them.
- CRITICAL: The HTML section between <!-- START FOOTER --> and <!-- END FOOTER --> is the footer. It must ALWAYS be present and completely intact in your output. Never remove, truncate, or modify it under any circumstances.
- Only replace these plain text placeholders:
    BODY_CONTENT_HERE → email body copy written from the ticket description
    BUTTON_TEXT       → a relevant CTA based on ticket context
    JOB_NUMBER_HERE   → the job number, or remove the line if blank
- Write body copy in a professional tone appropriate to the template and product.
- If the description says TBD or TBC, write appropriate placeholder copy based on the subject line and product.`;

// ─────────────────────────────────────────────
// Helper: Monday GraphQL
// ─────────────────────────────────────────────
async function mondayQuery(query, variables = {}) {
  const res = await fetch(MONDAY_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization:  process.env.MONDAY_API_TOKEN,
      "API-Version":  "2024-01",
    },
    body: JSON.stringify({ query, variables }),
  });
  const data = await res.json();
  if (data.errors) throw new Error(data.errors.map(e => e.message).join(", "));
  return data.data;
}

// ─────────────────────────────────────────────
// Helper: normalize Monday column_values
// ─────────────────────────────────────────────
function normalizeTicket(item) {
  const cols = {};
  for (const cv of item.column_values) cols[cv.id] = cv.text ?? cv.value ?? "";
  return {
    id:          item.id,
    name:        item.name,
    status:      cols["status"]        || "",
    description: cols["long_text7"]    || "",
    subjectLine: cols["text86"]        || "",
    jobNumber:   cols["formula"]       || "",
    sendDate:    cols["date4"]         || "",
    contentDue:  cols["date_mkx4g1zc"] || "",
    type:        cols["dropdown3"]     || "",
    category:    cols["status_1"]      || "",
    product:     cols["dropdown2"]     || "",
    requestor:   cols["person"]        || "",
    instructions: cols["long_text_mm47njms"] || "",  // Instructions column
    hasFiles:    !!cols["files"],
  };
}

// ─────────────────────────────────────────────
// Helper: fetch a single item by ID
// ─────────────────────────────────────────────
async function fetchItemById(itemId) {
  const data = await mondayQuery(`
    query GetItem($itemId: ID!) {
      items(ids: [$itemId]) {
        id name
        column_values(ids: [
          "status","long_text7","text86","formula",
          "date4","date_mkx4g1zc","dropdown3",
          "status_1","dropdown2","person","files","long_text_mm47njms"
        ]) { id text value }
      }
    }
  `, { itemId });
  const item = data?.items?.[0];
  if (!item) throw new Error(`Item ${itemId} not found.`);
  return normalizeTicket(item);
}

// ─────────────────────────────────────────────
// Helper: parse Instructions column into a variable map
// Format supported:
//   PreheaderText: "some value"
//   PreheaderText: some value (no quotes)
//   PreheaderText:
//   "some value"   (value on next line)
// ─────────────────────────────────────────────
function parseInstructions(instructions) {
  const vars = {};
  if (!instructions) return vars;

  const VARIABLE_NAMES = [
    "PreheaderText","BodyText","PrimaryLink","PrimaryText",
    "SecondaryLink","SecondText","TertiaryLink","TertiaryText"
  ];

  // Build a regex that matches "VariableName: value" or "VariableName:\n value"
  const pattern = new RegExp(
    `(${VARIABLE_NAMES.join("|")})\\s*:\\s*"?([^"\\n]*(?:\\n(?!(?:${VARIABLE_NAMES.join("|")})\\s*:)[^\\n]*)*)"?`,
    "gi"
  );

  let match;
  while ((match = pattern.exec(instructions)) !== null) {
    const key = match[1].trim();
    const val = match[2].trim().replace(/^"|"$/g, "");
    if (key && val) vars[key] = val;
  }

  return vars;
}

// ─────────────────────────────────────────────
// Helper: replace {{Variable}} tokens in HTML
// Only replaces tokens that are NOT Pardot tags
// (Pardot tags are {{Recipient.*}}, {{{...}}}, etc.)
// ─────────────────────────────────────────────
const CONTENT_VARIABLES = [
  "PreheaderText","BodyText","PrimaryLink","PrimaryText",
  "SecondaryLink","SecondText","TertiaryLink","TertiaryText",
  "Subject","JobNumber"
];

function applyVariables(html, vars) {
  let out = html;
  for (const key of CONTENT_VARIABLES) {
    if (vars[key] !== undefined) {
      const regex = new RegExp(`\\{\\{${key}\\}\\}`, "g");
      out = out.replace(regex, vars[key]);
    }
  }
  // Also replace legacy plain-text placeholders
  if (vars["BodyText"])   out = out.replace(/BODY_CONTENT_HERE/g, vars["BodyText"]);
  if (vars["JobNumber"])  out = out.replace(/JOB_NUMBER_HERE/g, vars["JobNumber"])
                                   .replace(/JOB_NUMBER(?!_)/g, vars["JobNumber"]);
  return out;
}
function extractFooter(html) {
  const match = html.match(/<!--\s*START FOOTER\s*-->([\s\S]*?)<!--\s*END FOOTER\s*-->/i);
  return match ? match[0] : null;
}

function hasFooter(html) {
  return /<!--\s*START FOOTER\s*-->/i.test(html) && /<!--\s*END FOOTER\s*-->/i.test(html);
}

function reattachFooter(html, footer) {
  // Insert footer before closing </table> of the main wrapper
  const insertPoint = html.lastIndexOf("</tbody></table></div>");
  if (insertPoint !== -1) {
    return html.slice(0, insertPoint) + "\n" + footer + "\n" + html.slice(insertPoint);
  }
  // Fallback: insert before </body>
  return html.replace("</body>", footer + "\n</body>");
}

// ─────────────────────────────────────────────
// Helper: generate HTML via Claude
// ─────────────────────────────────────────────
async function generateHTML(ticket, templateName) {
  const templateHtml = TEMPLATES[templateName];
  if (!templateHtml) throw new Error(`Unknown template: ${templateName}`);

  const originalFooter = extractFooter(templateHtml);

  // Step 1: Parse Instructions column into variable map
  const vars = parseInstructions(ticket.instructions);

  // Step 2: Inject Subject and JobNumber from their own columns
  vars["Subject"]    = ticket.subjectLine || ticket.name || "";
  vars["JobNumber"]  = ticket.jobNumber || "";

  // Step 3: Apply all known variables directly — no AI needed for these
  let html = applyVariables(templateHtml, vars);

  // Step 4: Check which content placeholders still need filling
  const stillMissing = [];
  if (/BODY_CONTENT_HERE/.test(html) || /\{\{BodyText\}\}/.test(html)) stillMissing.push("BodyText");
  if (/\{\{PreheaderText\}\}/.test(html)) stillMissing.push("PreheaderText");
  if (/BUTTON_TEXT/.test(html))           stillMissing.push("ButtonText");

  // Step 5: Only call Claude if there are unfilled placeholders
  if (stillMissing.length > 0) {
    const message = await anthropic.messages.create({
      model:      "claude-haiku-4-5-20251001",
      max_tokens: 8000,
      system:     AI_SYSTEM_PROMPT,
      messages: [{
        role:    "user",
        content: `The following placeholders still need content: ${stillMissing.join(", ")}

Ticket data:
Name: ${ticket.name}
Template: ${templateName}
Subject Line: ${vars["Subject"]}
Job Number: ${vars["JobNumber"]}
Description: ${ticket.description || "No description provided."}
Category: ${ticket.category || ""}
Product: ${ticket.product || ""}

HTML TEMPLATE (partially populated — only fill the remaining placeholders listed above):
${html}`,
      }],
    });

    html = message.content.find(b => b.type === "text")?.text ?? html;
    html = html.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
  }

  // Step 6: Footer enforcement
  if (originalFooter && !hasFooter(html)) {
    console.warn(`Footer missing for "${ticket.name}" — reattaching.`);
    html = reattachFooter(html, originalFooter);
  }

  return html;
}

// ─────────────────────────────────────────────
// Helper: upload HTML file to Monday item
// ─────────────────────────────────────────────
async function uploadToMonday(itemId, fileName, html) {
  const query = `
    mutation AddFileToColumn($itemId: ID!, $columnId: String!, $file: File!) {
      add_file_to_column(item_id: $itemId, column_id: $columnId, file: $file) {
        id name url
      }
    }
  `;

  const form = new FormData();
  form.append("query", query);
  form.append("variables", JSON.stringify({ itemId, columnId: "files", file: null }));
  form.append("map", JSON.stringify({ "0": ["variables.file"] }));

  const buf = Buffer.from(html, "utf-8");
  form.append("0", buf, { filename: fileName, contentType: "text/html", knownLength: buf.length });

  const res = await fetch(MONDAY_FILE_URL, {
    method:  "POST",
    headers: { Authorization: process.env.MONDAY_API_TOKEN, "API-Version": "2024-01", ...form.getHeaders() },
    body:    form,
  });

  const data = await res.json();
  if (data.errors) throw new Error(data.errors.map(e => e.message).join(", "));
  return data?.data?.add_file_to_column;
}

// ═════════════════════════════════════════════
// ROUTES
// ═════════════════════════════════════════════

// Root + health
app.get("/",       (_, res) => res.json({ status: "ok", service: "Monday Email Generator" }));
app.get("/health", (_, res) => res.json({ status: "ok" }));

// GET /api/tickets — fetch all groups, filter for New Requests
app.get("/api/tickets", async (req, res) => {
  try {
    const data = await mondayQuery(`
      query GetGroups($boardId: ID!) {
        boards(ids: [$boardId]) {
          groups {
            id title
            items_page(limit: 50) {
              items {
                id name
                column_values(ids: [
                  "status","long_text7","text86","formula",
                  "date4","date_mkx4g1zc","dropdown3",
                  "status_1","dropdown2","person","files","long_text_mm47njms"
                ]) { id text value }
              }
            }
          }
        }
      }
    `, { boardId: BOARD_ID });

    const allGroups = data?.boards?.[0]?.groups ?? [];
    const matched   = allGroups.filter(g => g.title.toLowerCase().includes("new request"));
    const groups    = matched.length > 0 ? matched : allGroups;
    const tickets   = groups.flatMap(g => g.items_page?.items ?? []).map(normalizeTicket);

    res.json({ tickets });
  } catch (err) {
    console.error("GET /api/tickets:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/generate — generate HTML from ticket + template
app.post("/api/generate", async (req, res) => {
  const { ticket, templateName } = req.body;
  if (!ticket || !templateName) return res.status(400).json({ error: "Missing ticket or templateName" });
  try {
    const html = await generateHTML(ticket, templateName);
    res.json({ html });
  } catch (err) {
    console.error("POST /api/generate:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/upload — upload HTML file to Monday ticket
app.post("/api/upload", async (req, res) => {
  const { itemId, fileName, html } = req.body;
  if (!itemId || !fileName || !html) return res.status(400).json({ error: "Missing itemId, fileName, or html" });
  try {
    const asset = await uploadToMonday(itemId, fileName, html);
    res.json({ success: true, assetId: asset?.id, fileName: asset?.name, url: asset?.url });
  } catch (err) {
    console.error("POST /api/upload:", err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// POST /api/webhook — Monday fires this when a new item is created
app.post("/api/webhook", async (req, res) => {
  const { challenge, event } = req.body;

  // Monday sends a challenge on first setup — must echo it back
  if (challenge) return res.json({ challenge });

  // Only handle item_created events
  if (event?.type !== "create_pulse") {
    return res.json({ status: "ignored" });
  }

  const itemId = String(event.pulseId);
  console.log(`Webhook: new item created — ID ${itemId}`);

  // Respond immediately so Monday doesn't retry
  res.json({ status: "received" });

  // Process asynchronously
  try {
    const ticket = await fetchItemById(itemId);
    console.log(`Processing ticket: ${ticket.name}`);

    // Pick template — from Template column, fallback to CPACOM General
    const templateName = TEMPLATES[ticket.template] ? ticket.template : "CPACOM General";
    console.log(`Using template: ${templateName}`);

    const html     = await generateHTML(ticket, templateName);
    const fileName = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}.html`;
    const asset    = await uploadToMonday(itemId, fileName, html);

    console.log(`✅ Uploaded ${fileName} to item ${itemId} — asset ID: ${asset?.id}`);
  } catch (err) {
    console.error(`Webhook processing error for item ${itemId}:`, err.message);
  }
});

// ─────────────────────────────────────────────
// Serve frontend UI
// ─────────────────────────────────────────────
app.get("/app", (_, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅  Server running on port ${PORT}`));
