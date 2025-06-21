package me.metallicgoat.datacombiner;

import java.util.*;


// ENUM name is lab data versions
// ENUM oaraometers are valyes that may appear in index files

// Convert slashes, spaces, commas, and dashes to underscore
// Case is irrelevant, lower preferred

public enum Param {
  ag("silver"),
  al("aluminum"),
  alkalinity_as_caco3("alkalinity"),
  as("arsenic"),
  b("boron"),
  ba("barium"),
  be("beryllium"),
  bi("bismuth"),
  bromodichloromethane(),
  bromoform(),
  bromomethane(),
  bod("BOD5", "biochemical oxygen demand"),
  c_1_2_dichloroethylene("cis-1,2-Dichloroethylene"),
  c_1_3_dichloropropylene("cis-1,3-Dichloropropylene"),
  ca("calcium"),
  cation_sum(),
  cd("cadmium"),
  cl("chloride", "chlorides"),
  co("cobalt"),
  co3_as_caco3("carbonate"),
  cod("chemical oxygen demand"),
  colour__true_("colour"),
  conductivity(),
  cr("chromium"),
  cu("copper"),
  dibromochloromethane(),
  dichloromethane(),
  doc("dissolved organic carbon"),
  ethylbenzene("ethyl benzene"),
  ethylene_dibromide(),
  fe("iron"),
  hardness_as_caco3("hardness"),
  hco3_as_caco3("bicarbonate"),
  ion_balance("ion ratio"),
  k("potassium"),
  m_p_xylene(),
  methyl_ethyl_ketone__mek_("methyl ethyl ketone", "mek"),
  methyl_isobutyl_ketone__mibk_("methyl isobutyl ketone", "mibk"),
  methyl_tert_butyl_ether__mtbe_("methyl-t-butyl ether", "mtbe"),
  mg("magnesium"),
  mn("manganese"),
  mo("molybdenum"),
  monochlorobenzene("chlorobenzene"),
  na("sodium"),
  ni("nickel"),
  n_nh3("ammonia", "ammonia: total"),
  n_no2("nitrite"),
  n_no3("nitrate"),
  o_po4("phosphate-ortho", "phosphate - ortho", "orthophosphate", "ortho-phosphate", "phosphate-O"),
  o_xylene(),
  p("phosphorus"),
  pb("lead"),
  ph(),
  phenols(),
  sb("antimony"),
  se("selenium"),
  si("silicon", "silica"),
  sn("tin"),
  so4("sulphate", "sulphates"),
  sr("strontium"),
  styrene(),
  t_1_2_dichloroethylene("trans-1,2-Dichloroethylene"),
  t_1_3_dichloropropylene("trans-1,3-Dichloropropylene"),
  tds__cond___calc_("total dissolved solids"),
  tss("total suspended solids"),
  ti("titanium"),
  tl("thallium"),
  toluene(),
  toluene_d8(),
  total_kjeldahl_nitrogen(),
  trichloroethylene(),
  turbidity(),
  u("uranium"),
  v("vanadium"),
  xylene_total("xylenes - total"),
  zn("zinc");

  private final Set<String> alternatives;

  Param(String... alternatives) {
    this.alternatives = new HashSet<>();
    for (String alt : alternatives) {
      this.alternatives.add(standardizeParam( alt.toLowerCase()));
    }
  }

  private static final Map<String, Param> LOOKUP_MAP = new HashMap<>();

  static {
    for (Param param : values()) {
      // Add enum name itself
      LOOKUP_MAP.put(standardizeParam(param.name()), param);
      // Add all alternative names
      for (String alt : param.alternatives) {
        LOOKUP_MAP.put(standardizeParam(alt), param);
      }
    }
  }

  public static Param getParamByAlternativeName(String name) {
    if (name == null) return null;
    return LOOKUP_MAP.get(standardizeParam(name));
  }


  public static String getStandardName(String param) {
    final Param p = getParamByAlternativeName(param);

    if (p != null) {
      return standardizeParam(p.name());

    } else {
      // If no match found, return the original param as is
      return standardizeParam(param.toLowerCase());
    }
  }

  public static String standardizeParam(String param) {
    return param.toLowerCase().trim().replace("-", "_").replace(" ", "_").replace(",", "_").replace("/", "_").replace("(", "_").replace(")", "_");
  }
}
