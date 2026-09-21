// Render react-icons to PNGs for the deck
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

const fa = require("react-icons/fa");
const md = require("react-icons/md");
const gi = require("react-icons/gi");
const tb = require("react-icons/tb");
const lu = require("react-icons/lu");
const bi = require("react-icons/bi");
const pi = require("react-icons/pi");

const OUT = path.join(__dirname, "icons");
fs.mkdirSync(OUT, { recursive: true });

// name -> [IconComponent, color]
const BURGUNDY = "#8C2B35";
const BLUE = "#2A78D6";
const WHITE = "#FFFFFF";
const INK = "#232323";

const icons = {
  // concept icons
  "syringe-white": [fa.FaSyringe, WHITE],
  "syringe-burgundy": [fa.FaSyringe, BURGUNDY],
  "virus-white": [fa.FaViruses, WHITE],
  "virus-burgundy": [fa.FaViruses, BURGUNDY],
  "pills-white": [fa.FaPills, WHITE],
  "pills-burgundy": [fa.FaPills, BURGUNDY],
  "ultrasound-white": [md.MdMonitorHeart, WHITE],
  "ultrasound-burgundy": [md.MdMonitorHeart, BURGUNDY],
  "liver-white": [gi.GiLiver, WHITE],
  "liver-burgundy": [gi.GiLiver, BURGUNDY],
  "liver-blue": [gi.GiLiver, BLUE],
  "search-white": [fa.FaSearch, WHITE],
  "search-blue": [fa.FaSearch, BLUE],
  "search-burgundy": [fa.FaSearch, BURGUNDY],
  "people-white": [md.MdGroups, WHITE],
  "people-blue": [md.MdGroups, BLUE],
  "people-burgundy": [md.MdGroups, BURGUNDY],
  "chart-down-white": [fa.FaChartLine, WHITE],
  "shield-white": [fa.FaShieldAlt, WHITE],
  "shield-burgundy": [fa.FaShieldAlt, BURGUNDY],
  "shield-blue": [fa.FaShieldAlt, BLUE],
  "hospital-white": [fa.FaHospital, WHITE],
  "flag-white": [fa.FaFlagCheckered, WHITE],
  "flag-blue": [fa.FaFlagCheckered, BLUE],
  "target-white": [tb.TbTargetArrow, WHITE],
  "target-blue": [tb.TbTargetArrow, BLUE],
  "target-burgundy": [tb.TbTargetArrow, BURGUNDY],
  "euro-white": [fa.FaEuroSign, WHITE],
  "euro-blue": [fa.FaEuroSign, BLUE],
  "wine-white": [bi.BiWine, WHITE],
  "wine-blue": [bi.BiWine, BLUE],
  "wine-burgundy": [bi.BiWine, BURGUNDY],
  "scale-white": [md.MdMonitorWeight, WHITE],
  "scale-blue": [md.MdMonitorWeight, BLUE],
  "dna-white": [fa.FaDna, WHITE],
  "dna-blue": [fa.FaDna, BLUE],
  "ai-white": [pi.PiRobotBold, WHITE],
  "ai-blue": [pi.PiRobotBold, BLUE],
  "blood-white": [fa.FaTint, WHITE],
  "blood-blue": [fa.FaTint, BLUE],
  "blood-burgundy": [fa.FaTint, BURGUNDY],
  "mri-white": [md.MdOutlineMedicalServices, WHITE],
  "mri-blue": [md.MdOutlineMedicalServices, BLUE],
  "baby-white": [fa.FaBaby, WHITE],
  "baby-burgundy": [fa.FaBaby, BURGUNDY],
  "gov-white": [md.MdAccountBalance, WHITE],
  "gov-burgundy": [md.MdAccountBalance, BURGUNDY],
  "gov-blue": [md.MdAccountBalance, BLUE],
  "link-white": [fa.FaLink, WHITE],
  "link-burgundy": [fa.FaLink, BURGUNDY],
  "calendar-white": [fa.FaCalendarCheck, WHITE],
  "calendar-blue": [fa.FaCalendarCheck, BLUE],
  "calendar-burgundy": [fa.FaCalendarCheck, BURGUNDY],
  "check-white": [fa.FaCheckCircle, WHITE],
  "check-green": [fa.FaCheckCircle, "#0CA30C"],
  "cross-red": [md.MdCancel, "#D03B3B"],
  "warn-amber": [fa.FaExclamationTriangle, "#C98500"],
  "arrow-right-white": [fa.FaArrowRight, WHITE],
  "arrow-right-ink": [fa.FaArrowRight, INK],
  "arrow-down-green": [fa.FaArrowDown, "#0CA30C"],
  "trophy-white": [fa.FaTrophy, WHITE],
  "trophy-gold": [fa.FaTrophy, "#EDA100"],
  "globe-white": [fa.FaGlobeEurope, WHITE],
  "globe-blue": [fa.FaGlobeEurope, BLUE],
  "handshake-white": [fa.FaHandshake, WHITE],
  "handshake-blue": [fa.FaHandshake, BLUE],
  "microscope-white": [fa.FaMicroscope, WHITE],
  "microscope-blue": [fa.FaMicroscope, BLUE],
  "roadmap-white": [lu.LuMap, WHITE],
  "stetho-white": [fa.FaStethoscope, WHITE],
  "stetho-blue": [fa.FaStethoscope, BLUE],
  "stetho-burgundy": [fa.FaStethoscope, BURGUNDY],
  "book-white": [fa.FaBookOpen, WHITE],
  "book-burgundy": [fa.FaBookOpen, BURGUNDY],
  "lightbulb-white": [fa.FaLightbulb, WHITE],
  "lightbulb-gold": [fa.FaLightbulb, "#EDA100"],
};

(async () => {
  for (const [name, [Icon, color]] of Object.entries(icons)) {
    if (!Icon) { console.error("MISSING ICON COMPONENT:", name); continue; }
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Icon, { color, size: 256 })
    );
    await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png().toFile(path.join(OUT, `${name}.png`));
  }
  console.log("done:", fs.readdirSync(OUT).length, "icons");
})();
