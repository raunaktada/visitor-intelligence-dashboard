export type TabId = 'customers' | 'industrial' | 'healthcare' | 'cpg' | 'other' | 'under1b';

// Add known TADA customer domains here to populate the Customers tab
export const CUSTOMER_DOMAINS: Set<string> = new Set([
  // e.g. 'caterpillar.com', 'osfhealthcare.org'
]);

// High-precision domain overrides (checked before keyword matching)
const DOMAIN_OVERRIDES: Record<string, TabId> = {
  // === HEALTHCARE ===
  'gehealthcare.com': 'healthcare', 'abbott.com': 'healthcare',
  'stryker.com': 'healthcare', 'medtronic.com': 'healthcare',
  'si-bone.com': 'healthcare', 'edwards.com': 'healthcare',
  'cpousa.com': 'healthcare', 'pfizer.com': 'healthcare',
  'abbvie.com': 'healthcare', 'drreddys.com': 'healthcare',
  'freseniusmedicalcare.com': 'healthcare', 'davita.com': 'healthcare',
  'uhc.com': 'healthcare', 'unitedhealthgroup.com': 'healthcare',
  'optum.com': 'healthcare', 'hcahealthcare.com': 'healthcare',
  'hcamidwest.com': 'healthcare', 'sanofi.com': 'healthcare',
  'parexel.com': 'healthcare', 'hologic.com': 'healthcare',
  'natera.com': 'healthcare', 'steris.com': 'healthcare',
  'osfhealthcare.org': 'healthcare', 'advocatehealth.com': 'healthcare',
  'bannerhealth.com': 'healthcare', 'christianacare.org': 'healthcare',
  'nm.org': 'healthcare', 'ochsner.org': 'healthcare',
  'carle.org': 'healthcare', 'crouse.org': 'healthcare',
  'froedtert.com': 'healthcare', 'lifestance.com': 'healthcare',
  'adventhealth.com': 'healthcare', 'nhs.net': 'healthcare',
  'mobileaspects.com': 'healthcare', 'senderrarx.com': 'healthcare',
  'texomamedicalcenter.net': 'healthcare', 'endologix.com': 'healthcare',
  'bronsonhealth.com': 'healthcare', 'choa.org': 'healthcare',
  'catholichealthli.org': 'healthcare', 'childrenscolorado.org': 'healthcare',
  'childrensmn.org': 'healthcare', 'chkd.org': 'healthcare',
  'bvhealthsystem.org': 'healthcare', 'encompasshealth.com': 'healthcare',
  'endeavorhealth.org': 'healthcare', 'endo.com': 'healthcare',
  'fmolhs.org': 'healthcare', 'hackensackmeridianhealth.org': 'healthcare',
  'hollyhillhospital.com': 'healthcare', 'iconplc.com': 'healthcare',
  'icumed.com': 'healthcare', 'jacksonhealth.org': 'healthcare',
  'hopkinsmedicine.org': 'healthcare', 'lifehealthcare.com.au': 'healthcare',
  'mdanderson.org': 'healthcare', 'marywashingtonhealthcare.com': 'healthcare',
  'maryhaven.com': 'healthcare', 'nicklaushealth.org': 'healthcare',
  'nyp.org': 'healthcare', 'onco360.com': 'healthcare',
  'orthopediatrics.com': 'healthcare', 'outcomes.com': 'healthcare',
  'pancarefl.org': 'healthcare', 'parashospitals.com': 'healthcare',
  'pdrx.com': 'healthcare', 'pinerest.org': 'healthcare',
  'providence.org': 'healthcare', 'radpartners.com': 'healthcare',
  'rmlspecialtyhospital.org': 'healthcare', 'sca.health': 'healthcare',
  'spectrumhealthsystems.org': 'healthcare', 'springfieldclinic.com': 'healthcare',
  'starkey.com': 'healthcare', 'sutterhealth.org': 'healthcare',
  'texaschildrens.org': 'healthcare', 'towerhealth.org': 'healthcare',
  'universityhealth.com': 'healthcare', 'vumc.org': 'healthcare',
  'ynhh.org': 'healthcare', 'atlanticare.org': 'healthcare',
  'elevancehealth.com': 'healthcare', 'barringtonortho.com': 'healthcare',
  'bioplusrx.com': 'healthcare', 'conehealth.com': 'healthcare',
  'convergentdental.com': 'healthcare', 'hillcrest.com': 'healthcare',
  'hmr.net': 'healthcare', 'hurleymc.com': 'healthcare',
  'hyperbaricmedicalservices.com': 'healthcare', 'lachmanconsultants.com': 'healthcare',
  'mchonline.org': 'healthcare', 'moffitt.org': 'healthcare',
  'oakstreethealth.com': 'healthcare', 'ocli.net': 'healthcare',
  'owensborohealth.org': 'healthcare', 'paloshealth.com': 'healthcare',
  'piedmont.org': 'healthcare', 'rhrhcenter.com': 'healthcare',
  'deeyesurgeons.com': 'healthcare', 'dodgecitysmiles.com': 'healthcare',
  'bostonprosthodontics.com': 'healthcare', 'charlestonradiologists.com': 'healthcare',
  'choicesbhc.com': 'healthcare', 'chconline.org': 'healthcare',
  'csl.com': 'healthcare', 'copalacupuncture.com': 'healthcare',
  'chail.org': 'healthcare', 'christushealth.org': 'healthcare',
  'brightlightimaging.com': 'healthcare', 'bms.com': 'healthcare',
  'angelsgracehospice.com': 'healthcare', 'abouthealth.com': 'healthcare',
  'truenorthwilderness.com': 'healthcare', 'sunnybrookhealthstore.com': 'healthcare',
  'healthnetworklabs.com': 'healthcare', 'hhsil.com': 'healthcare',
  'cityblueimaging.com': 'healthcare', 'cvshealth.com': 'healthcare',
  'walgreens.com': 'healthcare', 'brownhealth.org': 'healthcare',
  'baptisthealth.com': 'healthcare', 'brightstarcare.com': 'healthcare',
  'osfhealthcare.org': 'healthcare', 'carolina-health.com': 'healthcare',
  'hackensackmeridianhealth.org': 'healthcare', 'blickcenter.org': 'healthcare',

  // === INDUSTRIAL MANUFACTURERS ===
  'caterpillar.com': 'industrial', 'oshkoshcorp.com': 'industrial',
  'cummins.com': 'industrial', 'boeing.com': 'industrial',
  'honeywell.com': 'industrial', 'ge.com': 'industrial',
  'gm.com': 'industrial', 'collinsaerospace.com': 'industrial',
  'gdmissionsystems.com': 'industrial', 'wabteccorp.com': 'industrial',
  'parker.com': 'industrial', 'eaton.com': 'industrial',
  'cnh.com': 'industrial', 'kuka.com': 'industrial',
  'daimlertruck.com': 'industrial', 'cognex.com': 'industrial',
  'international.com': 'industrial', 'lennox.com': 'industrial',
  'warn.com': 'industrial', 'itw.com': 'industrial',
  'janicki.com': 'industrial', 'aerotech.com': 'industrial',
  'konecranes.com': 'industrial', 'mustangcat.com': 'industrial',
  'miltoncat.com': 'industrial', 'etnyre.com': 'industrial',
  'nordson.com': 'industrial', 'panduit.com': 'industrial',
  'oregontool.com': 'industrial', 'regalrexnord.com': 'industrial',
  'moog.com': 'industrial', 'hillenbrand.com': 'industrial',
  'alliedsystems.com': 'industrial', 'atlaselectric.net': 'industrial',
  'camco.net': 'industrial', 'clarios.com': 'industrial',
  'clarage.com': 'industrial', 'chemtreat.com': 'industrial',
  'kcindustries.com': 'industrial', 'draper.com': 'industrial',
  'tlgaerospace.com': 'industrial',
  'micropulseinc.com': 'industrial',
  'itdprecision.com': 'industrial', 'envirotechservices.com': 'industrial',
  'unityusa.com': 'industrial', 'ductmate.com': 'industrial',
  'richindustriesinc.com': 'industrial', 'carolinacat.com': 'industrial',
  'armstrongfluidtechnology.com': 'industrial', 'anchorinc.com': 'industrial',
  'yelvington.com': 'industrial', 'trimble.com': 'industrial',
  'stanleyblackanddecker.com': 'industrial', 'teradyne.com': 'industrial',
  'sensata.com': 'industrial', 'whirlpool.com': 'industrial',
  'johnsoncontrols.com': 'industrial', 'intevaproducts.com': 'industrial',
  'dana.com': 'industrial', 'douglasdynamics.com': 'industrial',
  'geappliances.com': 'industrial', 't3automation.com': 'industrial',
  'kamax.com': 'industrial', 'julianelectric.com': 'industrial',
  'kresscarrier.com': 'industrial', 'mahle.com': 'industrial',
  'steelscape.com': 'industrial', 'pbf.com': 'industrial',
  'ternium.com': 'industrial', 'spec-tech.com': 'industrial',
  'spectechind.com': 'industrial',

  // === CPG ===
  'target.com': 'cpg', 'walmart.com': 'cpg',
  'nestle.com': 'cpg', 'pepsico.com': 'cpg',
  'generalmills.com': 'cpg', 'cargill.com': 'cpg',
  'mars.com': 'cpg', 'pg.com': 'cpg',
  'unilever.com': 'cpg', 'heb.com': 'cpg',
  'segrocers.com': 'cpg', 'homedepot.com': 'cpg',
  'lowes.com': 'cpg', 'wayfair.com': 'cpg',
  'tjx.com': 'cpg', 'macys.com': 'cpg', 'macysinc.com': 'cpg',
  'nordstrom.com': 'cpg', 'kohls.com': 'cpg',
  'crocs.com': 'cpg', 'elcompanies.com': 'cpg',
  'colgatepalmolive.com': 'cpg', 'kenvue.com': 'cpg',
  'reckitt.com': 'cpg', 'flowersfoods.com': 'cpg',
  'staples.com': 'cpg', 'adm.com': 'cpg',
  'dudeproducts.com': 'cpg', 'bridgeandburn.com': 'cpg',
  'benekeith.com': 'cpg', 'coach.com': 'cpg',
  'crateandbarrel.com': 'cpg', 'footballfanatics.com': 'cpg',
  'papamurphys.com': 'cpg', 'petsmart.com': 'cpg',
  'rodanandfields.com': 'cpg', 'stellaandchewys.com': 'cpg',
  'worldmarket.com': 'cpg', 'bjs.com': 'cpg',
  'chicosfas.com': 'cpg', 'canadiantire.ca': 'cpg',
  'etsy.com': 'cpg', 'eyebuydirect.com': 'cpg',
  'lordandtaylor.com': 'cpg', 'shopmyexchange.com': 'cpg',
  'tiffany.com': 'cpg', 'instacart.com': 'cpg',
  'kwiktrip.com': 'cpg', 'mattel.com': 'cpg',
  'mrpricegroup.com': 'cpg', 'steinmart.com': 'cpg',
  'thefatquartergypsy.com': 'cpg', 'varsity.com': 'cpg',
  'ulta.com': 'cpg', 'converse.com': 'cpg',
  'apple.com': 'cpg', 'underarmour.com': 'cpg',
  'richs.com': 'cpg', 'goodyear.com': 'cpg',
  'daddario.com': 'cpg', 'flashfood.com': 'cpg',
  'allurebridals.com': 'cpg', 'circuitcity.com': 'cpg',
  'tileshop.com': 'cpg', 'viewrail.com': 'cpg',
  'klum.com': 'cpg', 'amgfoodservicesales.com': 'cpg',
  'dinshaws.co.in': 'cpg', 'generalimills.com': 'cpg',
};

const HEALTHCARE_KEYWORDS = [
  'health', 'hospital', 'clinic', 'medical', 'pharma', 'dental',
  'orthopedic', 'prosthetic', 'vision', 'rehab', 'therapy',
  'oncol', 'surgical', 'imaging', 'diagnostic', 'hospice',
  'nursing', 'healthcare', 'medtech', 'pharmacy', 'wellness',
  'radiolog', 'pediatric', 'ophthalmol', 'behavioral', 'naturopathic',
  'counseling', 'optom',
];

const INDUSTRIAL_KEYWORDS = [
  'industrial', 'aerospace', 'machinery', 'automotive', 'defense',
  'turbine', 'crane', 'equipment', 'engineering', 'electric',
  'semiconductor', 'automation', 'manufacturing', 'systems',
];

const CPG_KEYWORDS = [
  'grocery', 'retail', 'food', 'beverage', 'beauty', 'apparel',
  'fashion', 'cosmetic', 'consumer', 'supermarket', 'brewery',
  'restaurant', 'cafe', 'bakery',
];

export function classifyCompany(name: string, domain: string): TabId {
  const d = domain.toLowerCase();
  const n = name.toLowerCase();

  if (CUSTOMER_DOMAINS.has(d)) return 'customers';

  const override = DOMAIN_OVERRIDES[d];
  if (override) return override;

  // Keyword fallback
  if (HEALTHCARE_KEYWORDS.some(k => n.includes(k))) return 'healthcare';
  if (INDUSTRIAL_KEYWORDS.some(k => n.includes(k))) return 'industrial';
  if (CPG_KEYWORDS.some(k => n.includes(k))) return 'cpg';

  return 'other';
}
