"use client";

import {
  useEffect,
  useRef,
  useState,
  type FormEvent,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";
import { createPortal } from "react-dom";
import QRCode from "react-qr-code";
import Select, { type StylesConfig } from "react-select";
import { useReactToPrint } from "react-to-print";
import * as XLSX from "xlsx";
import { cities } from "../data/cities";
import {
  Area,
  AreaChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

type CityOption = {
  value: string;
  label: string;
};

const cityOptions: CityOption[] = cities.map((city) => ({
  value: city,
  label: city,
}));

const destinationSelectStyles: StylesConfig<CityOption, false> = {
  control: (provided, state) => ({
    ...provided,
    minHeight: "32px",
    borderColor: state.isFocused ? "#3b82f6" : "#d1d5db",
    borderRadius: "0.375rem",
    boxShadow: state.isFocused ? "0 0 0 2px #dbeafe" : "none",
    fontSize: "0.75rem",
    "&:hover": {
      borderColor: state.isFocused ? "#3b82f6" : "#d1d5db",
    },
  }),
  valueContainer: (provided) => ({
    ...provided,
    minWidth: 0,
    padding: "0 0.5rem",
  }),
  input: (provided) => ({
    ...provided,
    margin: 0,
    padding: 0,
  }),
  indicatorsContainer: (provided) => ({
    ...provided,
    height: "30px",
  }),
  dropdownIndicator: (provided) => ({
    ...provided,
    padding: "0 0.25rem",
  }),
  clearIndicator: (provided) => ({
    ...provided,
    padding: "0 0.25rem",
  }),
  menu: (provided) => ({
    ...provided,
    zIndex: 20,
  }),
};

type Docket = {
  docketNo: string;
  date: string;
  destination: string;
  courier: string;
  mode: string;
  weight: number;
  pieces: number;
  amount: number;
  taxableValue?: number;
  cgst?: number;
  sgst?: number;
  igst?: number;
};

type DocketForm = {
  docketNo: string;
  date: string;
  destination: string;
  courier: string;
  mode: string;
  weight: string;
  pieces: string;
  amount: string;
};

type Customer = {
  _id?: string;
  name: string;
  phone: string;
  address: string;
  state: string;
  gstin?: string;
  contactPerson?: string;
  email?: string;
  pan?: string;
  city?: string;
  stateCode?: string;
  pincode?: string;
};

type CourierPartner = {
  id: string;
  name: string;
  contactPerson: string;
  phone: string;
  notes: string;
};

type AuthCredentials = {
  email: string;
  password: string;
};

type LoggedInUser = {
  name: string;
  email: string;
  role: "admin" | "sub-user";
};

type CompanySettings = {
  name: string;
  gstin: string;
  phone: string;
  email: string;
  address: string;
  city: string;
  state: string;
  pincode: string;
  logo: string;
  signature: string;
  upiId: string;
  bankName: string;
  accountNumber: string;
  ifscCode: string;
  branchName: string;
};

type InvoiceType = "GST" | "Non-GST";
type InvoiceFilter = "all" | InvoiceType;

type SavedInvoice = {
  _id: string;
  invoiceNo: string;
  date: string;
  dueDate: string;
  customer: Customer;
  invoiceType?: InvoiceType;
  gstType: "CGST+SGST" | "IGST";
  consignments: Docket[];
  subtotal: number;
  cgstTotal: number;
  sgstTotal: number;
  igstTotal: number;
  totalAmount: number;
  status: string;
  paymentStatus: string;
  terms: string;
  gstInclusive?: boolean;
};

type MongoInvoice = {
  _id: string;
  invoiceNo: string;
  invoiceDate: string;
  dueDate: string;
  customerName: string;
  customerState: string;
  invoiceType: InvoiceType;
  gstType: "CGST+SGST" | "IGST";
  taxableValue: number;
  cgst: number;
  sgst: number;
  igst: number;
  totalAmount: number;
  status: string;
  paymentStatus: string;
  consignments: Docket[];
  customer: Customer;
  terms: string;
  gstInclusive: boolean;
};

type ActiveTab =
  | "dashboard"
  | "invoices"
  | "customers"
  | "couriers"
  | "reports"
  | "settings"
  | "team";

type TeamUser = {
  _id: string;
  name: string;
  email: string;
  role: "admin" | "sub-user";
  createdAt?: string;
};

const emptyDocketForm: DocketForm = {
  docketNo: "",
  date: "",
  destination: "",
  courier: "",
  mode: "",
  weight: "",
  pieces: "",
  amount: "",
};

const emptyCustomerForm: Customer = {
  name: "",
  phone: "",
  address: "",
  state: "Delhi",
  gstin: "",
  contactPerson: "",
  email: "",
  pan: "",
  city: "",
  stateCode: "",
  pincode: "",
};

const emptyCourierPartnerForm = {
  name: "",
  contactPerson: "",
  phone: "",
  notes: "",
};

const defaultCompanySettings: CompanySettings = {
  name: "AGC BILLING",
  gstin: "07AAAAA0000A1Z5",
  phone: "",
  email: "",
  address: "",
  city: "Delhi",
  state: "Delhi",
  pincode: "",
  logo: "",
  signature: "",
  upiId: "",
  bankName: "",
  accountNumber: "",
  ifscCode: "",
  branchName: "",
};

function getDefaultInvoiceTerms(city: string) {
  const jurisdiction = city.trim() || "[Your Company City]";
  return [
    "1. Payment due within 5 days of the invoice date.",
    "2. Goods once sold will not be taken back or exchanged.",
    "3. Interest @ 18% p.a. will be charged if payment is delayed.",
    `4. Subject to ${jurisdiction} jurisdiction only.`,
  ].join("\n");
}

const indianStates = [
  "Andhra Pradesh",
  "Arunachal Pradesh",
  "Assam",
  "Bihar",
  "Chhattisgarh",
  "Goa",
  "Gujarat",
  "Haryana",
  "Himachal Pradesh",
  "Jharkhand",
  "Karnataka",
  "Kerala",
  "Madhya Pradesh",
  "Maharashtra",
  "Manipur",
  "Meghalaya",
  "Mizoram",
  "Nagaland",
  "Odisha",
  "Punjab",
  "Rajasthan",
  "Sikkim",
  "Tamil Nadu",
  "Telangana",
  "Tripura",
  "Uttar Pradesh",
  "Uttarakhand",
  "West Bengal",
  "Andaman and Nicobar Islands",
  "Chandigarh",
  "Dadra and Nagar Haveli and Daman and Diu",
  "Delhi",
  "Jammu and Kashmir",
  "Ladakh",
  "Lakshadweep",
  "Puducherry",
];

const navigationItems: { id: ActiveTab; label: string }[] = [
  { id: "dashboard", label: "Dashboard" },
  { id: "invoices", label: "Invoices" },
  { id: "customers", label: "Customers" },
  { id: "couriers", label: "Courier Partners" },
  { id: "reports", label: "Reports" },
  { id: "settings", label: "Settings" },
  { id: "team", label: "Team" },
];

const currencyFormatter = new Intl.NumberFormat("en-IN", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

function formatCurrency(amount: number) {
  return `₹${currencyFormatter.format(amount)}`;
}

function calculateGstAmount(taxableValue: number) {
  return Math.round((taxableValue * 0.09 + Number.EPSILON) * 100) / 100;
}

function roundCurrencyAmount(amount: number) {
  return Math.round((amount + Number.EPSILON) * 100) / 100;
}

function calculateInclusiveDocketTax(
  finalAmount: number,
  gstType: SavedInvoice["gstType"],
  invoiceType: InvoiceType = "GST",
) {
  if (invoiceType === "Non-GST") {
    return { taxableValue: roundCurrencyAmount(finalAmount), cgst: 0, sgst: 0, igst: 0 };
  }

  const taxableValue = roundCurrencyAmount(finalAmount / 1.18);
  const totalGst = roundCurrencyAmount(finalAmount - taxableValue);
  const cgst = gstType === "CGST+SGST" ? roundCurrencyAmount(totalGst / 2) : 0;
  const sgst = gstType === "CGST+SGST" ? roundCurrencyAmount(totalGst - cgst) : 0;

  return {
    taxableValue,
    cgst,
    sgst,
    igst: gstType === "IGST" ? totalGst : 0,
  };
}

function calculateInclusiveInvoiceTotals(
  consignments: Docket[],
  gstType: SavedInvoice["gstType"],
  invoiceType: InvoiceType = "GST",
) {
  const totals = consignments.reduce(
    (result, docket) => {
      const tax = calculateInclusiveDocketTax(docket.amount, gstType, invoiceType);
      result.subtotal += tax.taxableValue;
      result.cgstTotal += tax.cgst;
      result.sgstTotal += tax.sgst;
      result.igstTotal += tax.igst;
      result.totalAmount += docket.amount;
      return result;
    },
    { subtotal: 0, cgstTotal: 0, sgstTotal: 0, igstTotal: 0, totalAmount: 0 },
  );

  return {
    subtotal: roundCurrencyAmount(totals.subtotal),
    cgstTotal: roundCurrencyAmount(totals.cgstTotal),
    sgstTotal: roundCurrencyAmount(totals.sgstTotal),
    igstTotal: roundCurrencyAmount(totals.igstTotal),
    totalAmount: roundCurrencyAmount(totals.totalAmount),
  };
}

function getInvoiceDateDefaults(now = new Date()) {
  const dateToInputValue = (date: Date) =>
    `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
      date.getDate(),
    ).padStart(2, "0")}`;
  const due = new Date(now);
  due.setDate(due.getDate() + 7);

  return {
    invoiceDate: dateToInputValue(now),
    dueDate: dateToInputValue(due),
  };
}

function getDateInputValue(value: string) {
  const isoDate = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (isoDate) return value;

  const parsedDate = new Date(value);
  if (Number.isNaN(parsedDate.getTime())) return getInvoiceDateDefaults().invoiceDate;

  return `${parsedDate.getFullYear()}-${String(parsedDate.getMonth() + 1).padStart(
    2,
    "0",
  )}-${String(parsedDate.getDate()).padStart(2, "0")}`;
}

function getStoredInvoiceDateInputValue(value: string): string | null {
  const isoDate = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (isoDate) {
    const year = Number(isoDate[1]);
    const month = Number(isoDate[2]);
    const day = Number(isoDate[3]);
    const parsedDate = new Date(year, month - 1, day);
    if (
      parsedDate.getFullYear() === year &&
      parsedDate.getMonth() === month - 1 &&
      parsedDate.getDate() === day
    ) {
      return value;
    }
    return null;
  }

  const parsedDate = new Date(value);
  if (Number.isNaN(parsedDate.getTime())) return null;
  return `${parsedDate.getFullYear()}-${String(parsedDate.getMonth() + 1).padStart(
    2,
    "0",
  )}-${String(parsedDate.getDate()).padStart(2, "0")}`;
}

function createUniqueInvoiceNumber(
  existingNumbers: string[],
  invoiceType: InvoiceType = "GST",
) {
  const existing = new Set(existingNumbers);
  const prefix = invoiceType === "GST" ? "GST" : "NGST";
  const firstNumber = invoiceType === "GST" ? 1024 : 5021;
  const seriesPattern = new RegExp(`^${prefix}-(\\d+)$`);
  let nextNumber = firstNumber;

  for (const existingNumber of existing) {
    const match = seriesPattern.exec(existingNumber);
    if (match) nextNumber = Math.max(nextNumber, Number(match[1]) + 1);
  }

  let candidate = `${prefix}-${nextNumber}`;
  while (existing.has(candidate)) {
    nextNumber += 1;
    candidate = `${prefix}-${nextNumber}`;
  }
  return candidate;
}

function isDocket(value: unknown): value is Docket {
  if (typeof value !== "object" || value === null) return false;
  const docket = value as Record<string, unknown>;

  return (
    typeof docket.docketNo === "string" &&
    typeof docket.date === "string" &&
    typeof docket.destination === "string" &&
    typeof docket.courier === "string" &&
    typeof docket.mode === "string" &&
    typeof docket.weight === "number" &&
    typeof docket.pieces === "number" &&
    typeof docket.amount === "number" &&
    (docket.taxableValue === undefined || typeof docket.taxableValue === "number") &&
    (docket.cgst === undefined || typeof docket.cgst === "number") &&
    (docket.sgst === undefined || typeof docket.sgst === "number") &&
    (docket.igst === undefined || typeof docket.igst === "number")
  );
}

function isCustomer(value: unknown): value is Customer {
  if (typeof value !== "object" || value === null) return false;
  const customer = value as Record<string, unknown>;

  return (
    typeof customer.name === "string" &&
    typeof customer.phone === "string" &&
    typeof customer.address === "string" &&
    typeof customer.state === "string" &&
    (customer.gstin === undefined || typeof customer.gstin === "string") &&
    (customer.contactPerson === undefined || typeof customer.contactPerson === "string") &&
    (customer.email === undefined || typeof customer.email === "string") &&
    (customer.pan === undefined || typeof customer.pan === "string") &&
    (customer.city === undefined || typeof customer.city === "string") &&
    (customer.stateCode === undefined || typeof customer.stateCode === "string") &&
    (customer.pincode === undefined || typeof customer.pincode === "string")
  );
}

function isCompanySettings(value: unknown): value is CompanySettings {
  if (typeof value !== "object" || value === null) return false;
  const settings = value as Record<string, unknown>;

  return [
    "name",
    "gstin",
    "phone",
    "email",
    "address",
    "city",
    "state",
    "pincode",
    "logo",
    "bankName",
    "accountNumber",
    "ifscCode",
    "branchName",
  ].every((field) => typeof settings[field] === "string") &&
    (settings.upiId === undefined || typeof settings.upiId === "string") &&
    (settings.signature === undefined || typeof settings.signature === "string");
}

function isCourierPartner(value: unknown): value is CourierPartner {
  if (typeof value !== "object" || value === null) return false;
  const partner = value as Record<string, unknown>;

  return (
    typeof partner.id === "string" &&
    typeof partner.name === "string" &&
    typeof partner.contactPerson === "string" &&
    typeof partner.phone === "string" &&
    typeof partner.notes === "string"
  );
}

function isAuthCredentials(value: unknown): value is AuthCredentials {
  if (typeof value !== "object" || value === null) return false;
  const credentials = value as Record<string, unknown>;

  return typeof credentials.email === "string" && typeof credentials.password === "string";
}

function isLoggedInUser(value: unknown): value is LoggedInUser {
  if (typeof value !== "object" || value === null) return false;
  const user = value as Record<string, unknown>;

  return (
    typeof user.name === "string" &&
    typeof user.email === "string" &&
    (user.role === "admin" || user.role === "sub-user")
  );
}

function isTeamUser(value: unknown): value is TeamUser {
  if (typeof value !== "object" || value === null) return false;
  const user = value as Record<string, unknown>;

  return (
    typeof user._id === "string" &&
    typeof user.name === "string" &&
    typeof user.email === "string" &&
    (user.role === "admin" || user.role === "sub-user") &&
    (user.createdAt === undefined || typeof user.createdAt === "string")
  );
}

function isMongoInvoice(value: unknown): value is MongoInvoice {
  if (typeof value !== "object" || value === null) return false;
  const invoice = value as Record<string, unknown>;

  return (
    typeof invoice._id === "string" &&
    typeof invoice.invoiceNo === "string" &&
    typeof invoice.invoiceDate === "string" &&
    typeof invoice.dueDate === "string" &&
    typeof invoice.customerName === "string" &&
    typeof invoice.customerState === "string" &&
    isCustomer(invoice.customer) &&
    (invoice.invoiceType === "GST" || invoice.invoiceType === "Non-GST") &&
    (invoice.gstType === "CGST+SGST" || invoice.gstType === "IGST") &&
    Array.isArray(invoice.consignments) &&
    invoice.consignments.every(isDocket) &&
    typeof invoice.taxableValue === "number" &&
    typeof invoice.cgst === "number" &&
    typeof invoice.sgst === "number" &&
    typeof invoice.igst === "number" &&
    typeof invoice.totalAmount === "number" &&
    typeof invoice.status === "string" &&
    typeof invoice.paymentStatus === "string" &&
    typeof invoice.terms === "string" &&
    typeof invoice.gstInclusive === "boolean"
  );
}

function toSavedInvoice(invoice: MongoInvoice): SavedInvoice {
  return {
    _id: invoice._id,
    invoiceNo: invoice.invoiceNo,
    date: invoice.invoiceDate,
    dueDate: invoice.dueDate,
    customer: invoice.customer,
    invoiceType: invoice.invoiceType,
    gstType: invoice.gstType,
    consignments: invoice.consignments,
    subtotal: invoice.taxableValue,
    cgstTotal: invoice.cgst,
    sgstTotal: invoice.sgst,
    igstTotal: invoice.igst,
    totalAmount: invoice.totalAmount,
    status: invoice.status,
    paymentStatus: invoice.paymentStatus,
    terms: invoice.terms,
    gstInclusive: invoice.gstInclusive,
  };
}

function NavIcon({ name }: { name: ActiveTab }) {
  const paths: Record<ActiveTab, React.ReactNode> = {
    dashboard: (
      <>
        <rect x="3" y="3" width="7" height="7" rx="1" />
        <rect x="14" y="3" width="7" height="7" rx="1" />
        <rect x="14" y="14" width="7" height="7" rx="1" />
        <rect x="3" y="14" width="7" height="7" rx="1" />
      </>
    ),
    invoices: (
      <>
        <path d="M7 3.75h7l4 4v12.5a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1v-15a1 1 0 0 1 1-1Z" />
        <path d="M14 4v4h4M9 13h6M9 17h6" />
      </>
    ),
    customers: (
      <>
        <circle cx="9" cy="8" r="3.25" />
        <path d="M3.5 20v-1.5a5.5 5.5 0 0 1 11 0V20ZM16 5.25a3.25 3.25 0 0 1 0 6.25M17 14a4.75 4.75 0 0 1 3.5 4.58V20" />
      </>
    ),
    couriers: (
      <>
        <path d="M3 6h11v12H3zM14 10h4l3 3v5h-7z" />
        <circle cx="7.5" cy="18" r="2" />
        <circle cx="17.5" cy="18" r="2" />
      </>
    ),
    reports: (
      <>
        <path d="M4 20V10M10 20V4M16 20v-7M22 20H2" />
      </>
    ),
    settings: (
      <>
        <circle cx="12" cy="12" r="3" />
        <path d="m19.4 15 .1.1 1.4 1.1-1.4 2.4-1.7-.6a7.8 7.8 0 0 1-1.5.9l-.3 1.8h-2.8l-.3-1.8a7.8 7.8 0 0 1-1.5-.9l-1.7.6-1.4-2.4 1.4-1.1a7 7 0 0 1 0-1.8l-1.4-1.1 1.4-2.4 1.7.6a7.8 7.8 0 0 1 1.5-.9l.3-1.8h2.8l.3 1.8a7.8 7.8 0 0 1 1.5.9l1.7-.6 1.4 2.4-1.4 1.1a7 7 0 0 1 0 1.8Z" />
      </>
    ),
    team: (
      <>
        <circle cx="9" cy="8" r="3" />
        <path d="M3.5 20v-1.5a5.5 5.5 0 0 1 11 0V20ZM16 5.25a3.25 3.25 0 0 1 0 6.25M17 14a4.75 4.75 0 0 1 3.5 4.58V20" />
      </>
    ),
  };

  return (
    <svg
      aria-hidden="true"
      className="h-5 w-5 shrink-0"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.7"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      {paths[name]}
    </svg>
  );
}

export default function Home() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [loggedInUser, setLoggedInUser] = useState<LoggedInUser | null>(null);
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [loginEmail, setLoginEmail] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [securityError, setSecurityError] = useState("");
  const [securitySuccess, setSecuritySuccess] = useState("");
  const [activeTab, setActiveTab] = useState<ActiveTab>("dashboard");
  const [invoiceFilter, setInvoiceFilter] = useState<InvoiceFilter>("all");
  const [teamUsers, setTeamUsers] = useState<TeamUser[]>([]);
  const [isLoadingTeamUsers, setIsLoadingTeamUsers] = useState(false);
  const [isAddingTeamUser, setIsAddingTeamUser] = useState(false);
  const [teamError, setTeamError] = useState("");
  const [teamForm, setTeamForm] = useState({
    name: "",
    email: "",
    password: "",
    role: "sub-user" as LoggedInUser["role"],
  });
  const [courierPartners, setCourierPartners] = useState<CourierPartner[]>([]);
  const [isAddingCourierPartner, setIsAddingCourierPartner] = useState(false);
  const [courierPartnerForm, setCourierPartnerForm] = useState(emptyCourierPartnerForm);
  const [courierError, setCourierError] = useState("");
  const [isCreatingInvoice, setIsCreatingInvoice] = useState(false);
  const [editingInvoiceId, setEditingInvoiceId] = useState<string | null>(null);
  const [form, setForm] = useState<DocketForm>(emptyDocketForm);
  const [dockets, setDockets] = useState<Docket[]>([]);
  const [isLoadingData, setIsLoadingData] = useState(true);
  const [isMigratingOldData, setIsMigratingOldData] = useState(false);
  const [savedInvoices, setSavedInvoices] = useState<SavedInvoice[]>([]);
  const [previewInvoiceNo, setPreviewInvoiceNo] = useState(() =>
    createUniqueInvoiceNumber([]),
  );
  const [selectedPrintInvoice, setSelectedPrintInvoice] = useState<SavedInvoice | null>(null);
  const printInvoiceRef = useRef<HTMLDivElement>(null);
  const [autoPrintInvoice, setAutoPrintInvoice] = useState(false);
  const [invoiceSuccess, setInvoiceSuccess] = useState("");
  const [invoiceError, setInvoiceError] = useState("");
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [customerForm, setCustomerForm] = useState<Customer>(emptyCustomerForm);
  const [isAddingCustomer, setIsAddingCustomer] = useState(false);
  const [customerError, setCustomerError] = useState("");
  const [companySettings, setCompanySettings] =
    useState<CompanySettings>(defaultCompanySettings);
  const [settingsError, setSettingsError] = useState("");
  const [settingsSuccess, setSettingsSuccess] = useState("");
  const [initialInvoiceDates] = useState(getInvoiceDateDefaults);
  const [invoiceDate, setInvoiceDate] = useState(initialInvoiceDates.invoiceDate);
  const [dueDate, setDueDate] = useState(initialInvoiceDates.dueDate);
  const [reportStartDate, setReportStartDate] = useState(() => {
    const [year, month] = getInvoiceDateDefaults().invoiceDate.split("-").map(Number);
    return `${year}-${String(month).padStart(2, "0")}-01`;
  });
  const [reportEndDate, setReportEndDate] = useState(
    () => getInvoiceDateDefaults().invoiceDate,
  );
  const [generatedReportRange, setGeneratedReportRange] = useState(() => {
    const endDate = getInvoiceDateDefaults().invoiceDate;
    const [year, month] = endDate.split("-").map(Number);
    return { startDate: `${year}-${String(month).padStart(2, "0")}-01`, endDate };
  });
  const [reportError, setReportError] = useState("");
  const [invoiceType, setInvoiceType] = useState<InvoiceType>("GST");
  const [selectedCustomer, setSelectedCustomer] = useState("");
  const [invoiceTerms, setInvoiceTerms] = useState(() =>
    getDefaultInvoiceTerms(defaultCompanySettings.city),
  );
  const [invoiceStatus, setInvoiceStatus] = useState("Issued");
  const [paymentStatus, setPaymentStatus] = useState("Unpaid");
  const filteredInvoices = savedInvoices.filter(
    (invoice) =>
      invoiceFilter === "all" || (invoice.invoiceType ?? "GST") === invoiceFilter,
  );
  const monthlyRevenueData = (() => {
    const now = new Date();
    const lastSixMonths = Array.from({ length: 6 }, (_, index) => {
      const monthDate = new Date(now.getFullYear(), now.getMonth() - 5 + index, 1);
      const year = monthDate.getFullYear();
      const month = monthDate.getMonth() + 1;
      return {
        month: new Intl.DateTimeFormat("en-IN", { month: "short" }).format(monthDate),
        monthKey: `${year}-${String(month).padStart(2, "0")}`,
        revenue: 0,
      };
    });
    const revenueByMonth = new Map(
      lastSixMonths.map((monthData) => [monthData.monthKey, monthData]),
    );

    for (const invoice of savedInvoices) {
      const invoiceDate = getStoredInvoiceDateInputValue(invoice.date);
      if (!invoiceDate) continue;

      const monthKey = invoiceDate.slice(0, 7);
      const monthData = revenueByMonth.get(monthKey);
      if (monthData) monthData.revenue += invoice.totalAmount;
    }

    return lastSixMonths;
  })();
  const reportInvoices = savedInvoices.filter((invoice) => {
    const invoiceDate = getStoredInvoiceDateInputValue(invoice.date);
    return (
      invoiceDate !== null &&
      invoiceDate >= generatedReportRange.startDate &&
      invoiceDate <= generatedReportRange.endDate
    );
  });
  const reportTotals = reportInvoices.reduce(
    (totals, invoice) => {
      const invoiceType = invoice.invoiceType ?? "GST";
      totals.totalSales += invoice.totalAmount;
      if (invoiceType === "Non-GST") {
        totals.nonGstSales += invoice.totalAmount;
      } else {
        totals.gstCollected += invoice.cgstTotal + invoice.sgstTotal + invoice.igstTotal;
      }
      return totals;
    },
    { totalSales: 0, gstCollected: 0, nonGstSales: 0 },
  );
  const handlePrint = useReactToPrint({
    contentRef: printInvoiceRef,
    documentTitle: "AGC_Invoice",
    onPrintError: (location, error) => {
      console.error(`Invoice print failed during ${location}.`, error);
      setInvoiceError("The invoice could not be opened for printing. Please try again.");
    },
  });

  const selectedCustomerRecord = customers.find(
    (customer) => customer.phone === selectedCustomer,
  ) ?? savedInvoices.find((invoice) => invoice._id === editingInvoiceId)?.customer;
  const isSameState =
    selectedCustomerRecord?.state.trim().toLocaleLowerCase() ===
    companySettings.state.trim().toLocaleLowerCase();
  const gstType = isSameState ? "CGST+SGST" : "IGST";
  const invoiceTotals = calculateInclusiveInvoiceTotals(dockets, gstType, invoiceType);
  const { subtotal, cgstTotal, sgstTotal, igstTotal } = invoiceTotals;
  const totalGst = cgstTotal + sgstTotal + igstTotal;
  const grandTotal = invoiceTotals.totalAmount;
  const printCompanySettings = { ...defaultCompanySettings, ...companySettings };
  const printConsignments = selectedPrintInvoice?.consignments ?? [];
  const printInvoiceType = selectedPrintInvoice?.invoiceType ?? "GST";
  const printGstType = selectedPrintInvoice?.gstType ?? "IGST";
  const printInclusiveTotals = calculateInclusiveInvoiceTotals(
    printConsignments,
    printGstType,
    printInvoiceType,
  );
  const printSubtotal =
    selectedPrintInvoice?.gstInclusive === true
      ? printInclusiveTotals.subtotal
      : selectedPrintInvoice?.subtotal ??
        printConsignments.reduce((total, docket) => total + docket.amount, 0);
  const printCgstTotal =
    selectedPrintInvoice?.gstInclusive === true
      ? printInclusiveTotals.cgstTotal
      : selectedPrintInvoice?.cgstTotal ??
        (printGstType === "CGST+SGST" ? calculateGstAmount(printSubtotal) : 0);
  const printSgstTotal =
    selectedPrintInvoice?.gstInclusive === true
      ? printInclusiveTotals.sgstTotal
      : selectedPrintInvoice?.sgstTotal ?? printCgstTotal;
  const printIgstTotal =
    selectedPrintInvoice?.gstInclusive === true
      ? printInclusiveTotals.igstTotal
      : selectedPrintInvoice?.igstTotal ??
        (printGstType === "IGST"
          ? roundCurrencyAmount(printSubtotal * 0.18)
          : 0);
  const printTotalAmount =
    selectedPrintInvoice?.gstInclusive === true
      ? printInclusiveTotals.totalAmount
      : selectedPrintInvoice?.totalAmount ??
        printSubtotal + printCgstTotal + printSgstTotal + printIgstTotal;
  const upiId = printCompanySettings.upiId?.trim();
  const upiPaymentUri = upiId
    ? `upi://pay?pa=${encodeURIComponent(upiId)}&pn=${encodeURIComponent(
        printCompanySettings.name,
      )}&cu=INR`
    : "";

  useEffect(() => {
    try {
      const storedCouriers = localStorage.getItem("agc_couriers");
      if (storedCouriers) {
        const parsedCouriers: unknown = JSON.parse(storedCouriers);
        if (Array.isArray(parsedCouriers)) {
          const validCouriers = parsedCouriers.filter(isCourierPartner);
          if (validCouriers.length !== parsedCouriers.length) {
            console.error("Some saved courier partners in localStorage have an invalid format.");
          }
          setCourierPartners(validCouriers);
        } else {
          console.error("Saved courier partner data in localStorage is not an array.");
        }
      }
      const storedCompanySettings = localStorage.getItem("agc_company_settings");
      if (storedCompanySettings) {
        const parsedSettings: unknown = JSON.parse(storedCompanySettings);
        if (isCompanySettings(parsedSettings)) {
          setCompanySettings({
            ...defaultCompanySettings,
            ...parsedSettings,
            signature: parsedSettings.signature ?? "",
            upiId: parsedSettings.upiId ?? "",
          });
        } else {
          console.error("Saved company settings in localStorage have an invalid format.");
        }
      }

    } catch (error) {
      console.error("Could not read saved account or company settings from localStorage.", error);
    }
  }, []);

  useEffect(() => {
    let active = true;

    async function restoreSession() {
      try {
        const response = await fetch("/api/auth/me");
        const result: unknown = await response.json();
        const user =
          typeof result === "object" && result !== null
            ? (result as Record<string, unknown>).user
            : null;

        if (!active) return;
        if (!response.ok || !isLoggedInUser(user)) {
          localStorage.removeItem("agc_user");
          setLoggedInUser(null);
          setIsLoggedIn(false);
          return;
        }

        localStorage.setItem("agc_user", JSON.stringify(user));
        setLoggedInUser(user);
        setLoginEmail(user.email);
        setIsLoggedIn(true);
      } catch (error) {
        if (!active) return;
        console.error("Could not verify the saved server session.", error);
        localStorage.removeItem("agc_user");
        setLoggedInUser(null);
        setIsLoggedIn(false);
      }
    }

    void restoreSession();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    let active = true;

    async function fetchCollection(url: string) {
      const response = await fetch(url);
      const data: unknown = await response.json();
      if (!response.ok) {
        const message =
          typeof data === "object" && data !== null &&
          typeof (data as Record<string, unknown>).error === "string"
            ? (data as Record<string, string>).error
            : `Request failed with status ${response.status}.`;
        throw new Error(message);
      }
      return data;
    }

    async function loadData() {
      if (!loggedInUser) {
        setIsLoadingData(false);
        setSavedInvoices([]);
        setCustomers([]);
        return;
      }
      setIsLoadingData(true);
      const [invoicesResult, customersResult] = await Promise.allSettled([
        fetchCollection("/api/invoices"),
        fetchCollection("/api/customers"),
      ]);

      if (!active) return;

      if (invoicesResult.status === "fulfilled") {
        const invoices = invoicesResult.value;
        if (Array.isArray(invoices) && invoices.every(isMongoInvoice)) {
          setSavedInvoices(invoices.map(toSavedInvoice));
        } else {
          console.error("Invoice API returned data in an invalid format.");
          setInvoiceError("Invoices could not be loaded because the server returned invalid data.");
        }
      } else {
        console.error("Could not load invoices from the API.", invoicesResult.reason);
        setInvoiceError("Invoices could not be loaded. Please refresh and try again.");
      }

      if (customersResult.status === "fulfilled") {
        const customersData = customersResult.value;
        if (
          Array.isArray(customersData) &&
          customersData.every(
            (customer) =>
              isCustomer(customer) &&
              typeof (customer as Record<string, unknown>)._id === "string",
          )
        ) {
          setCustomers(
            customersData.map((customer) => ({
              ...customer,
              gstin: customer.gstin ?? "",
              contactPerson: customer.contactPerson ?? "",
              email: customer.email ?? "",
              pan: customer.pan ?? "",
              city: customer.city ?? "",
              stateCode: customer.stateCode ?? "",
              pincode: customer.pincode ?? "",
            })),
          );
        } else {
          console.error("Customer API returned data in an invalid format.");
          setCustomerError("Customers could not be loaded because the server returned invalid data.");
        }
      } else {
        console.error("Could not load customers from the API.", customersResult.reason);
        setCustomerError("Customers could not be loaded. Please refresh and try again.");
      }

      setIsLoadingData(false);
    }

    void loadData();
    return () => {
      active = false;
    };
  }, [loggedInUser]);

  useEffect(() => {
    if (activeTab !== "team" || loggedInUser?.role !== "admin") return;
    let active = true;
    setIsLoadingTeamUsers(true);
    setTeamError("");

    async function loadTeamUsers() {
      try {
        const response = await fetch("/api/auth/users");
        const result: unknown = await response.json();
        if (!response.ok) {
          const message =
            typeof result === "object" &&
            result !== null &&
            typeof (result as Record<string, unknown>).error === "string"
              ? (result as Record<string, string>).error
              : "Could not load team users.";
          throw new Error(message);
        }
        if (
          !Array.isArray(result) ||
          !result.every(isTeamUser)
        ) {
          throw new Error("The users API returned invalid data.");
        }
        if (active) setTeamUsers(result);
      } catch (error) {
        if (!active) return;
        console.error("Could not load team users.", error);
        setTeamError(error instanceof Error ? error.message : "Could not load team users.");
      } finally {
        if (active) setIsLoadingTeamUsers(false);
      }
    }

    void loadTeamUsers();
    return () => {
      active = false;
    };
  }, [activeTab, loggedInUser]);

  useEffect(() => {
    if (autoPrintInvoice && selectedPrintInvoice) {
      handlePrint();
      setAutoPrintInvoice(false);
    }
  }, [autoPrintInvoice, handlePrint, selectedPrintInvoice]);

  useEffect(() => {
    if (!isAddingCustomer) return;

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") {
        setIsAddingCustomer(false);
        setCustomerError("");
      }
    }

    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [isAddingCustomer]);

  async function migrateOldDataToDB() {
    if (isMigratingOldData) return;

    setIsMigratingOldData(true);
    let migratedCustomers = 0;
    let migratedInvoices = 0;
    const migratedCustomerRecords: Customer[] = [];
    const migratedInvoiceRecords: SavedInvoice[] = [];

    try {
      const customerData = localStorage.getItem("agc_customers");
      const invoiceData = localStorage.getItem("agc_invoices");
      const parseStoredArray = (data: string | null, key: string): Record<string, unknown>[] => {
        if (!data) return [];

        const parsed: unknown = JSON.parse(data);
        if (
          !Array.isArray(parsed) ||
          !parsed.every(
            (item) => typeof item === "object" && item !== null && !Array.isArray(item),
          )
        ) {
          throw new Error(`Saved ${key} data is not a valid array of records.`);
        }
        return parsed as Record<string, unknown>[];
      };
      const customersToMigrate = parseStoredArray(customerData, "customer");
      const invoicesToMigrate = parseStoredArray(invoiceData, "invoice");
      const withoutIds = (record: Record<string, unknown>) => {
        const { id: _id, _id: _mongoId, ...payload } = record;
        return payload;
      };

      for (const customer of customersToMigrate) {
        const response = await fetch("/api/customers", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(withoutIds(customer)),
        });
        const result: unknown = await response.json();
        if (!response.ok) {
          const message =
            typeof result === "object" &&
            result !== null &&
            typeof (result as Record<string, unknown>).error === "string"
              ? (result as Record<string, string>).error
              : `Customer migration failed with status ${response.status}.`;
          throw new Error(message);
        }
        if (
          !isCustomer(result) ||
          typeof (result as Record<string, unknown>)._id !== "string"
        ) {
          throw new Error("The customer API returned invalid migration data.");
        }
        migratedCustomerRecords.push(result);
        migratedCustomers += 1;
      }

      for (const invoice of invoicesToMigrate) {
        const customer =
          typeof invoice.customer === "object" && invoice.customer !== null
            ? withoutIds(invoice.customer as Record<string, unknown>)
            : invoice.customer;
        const invoicePayload = {
          ...withoutIds(invoice),
          invoiceDate: invoice.invoiceDate ?? invoice.date,
          customerName: invoice.customerName ?? (customer as Customer | null)?.name,
          customerState: invoice.customerState ?? (customer as Customer | null)?.state,
          invoiceType: invoice.invoiceType ?? "GST",
          taxableValue: invoice.taxableValue ?? invoice.subtotal,
          cgst: invoice.cgst ?? invoice.cgstTotal,
          sgst: invoice.sgst ?? invoice.sgstTotal,
          igst: invoice.igst ?? invoice.igstTotal,
          customer,
          terms: typeof invoice.terms === "string" ? invoice.terms : "",
          gstInclusive:
            typeof invoice.gstInclusive === "boolean" ? invoice.gstInclusive : true,
        };
        const response = await fetch("/api/invoices", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(invoicePayload),
        });
        const result: unknown = await response.json();
        if (!response.ok) {
          const message =
            typeof result === "object" &&
            result !== null &&
            typeof (result as Record<string, unknown>).error === "string"
              ? (result as Record<string, string>).error
              : `Invoice migration failed with status ${response.status}.`;
          throw new Error(message);
        }
        if (!isMongoInvoice(result)) {
          throw new Error("The invoice API returned invalid migration data.");
        }
        migratedInvoiceRecords.push(toSavedInvoice(result));
        migratedInvoices += 1;
      }

      localStorage.removeItem("agc_customers");
      localStorage.removeItem("agc_invoices");
      setCustomers((currentCustomers) => [
        ...migratedCustomerRecords,
        ...currentCustomers,
      ]);
      setSavedInvoices((currentInvoices) => [
        ...migratedInvoiceRecords,
        ...currentInvoices,
      ]);
      window.alert(
        `Migration complete. ${migratedCustomers} customer(s) and ${migratedInvoices} invoice(s) copied to MongoDB.`,
      );
    } catch (error) {
      console.error("Could not migrate old LocalStorage records to MongoDB.", error);
      window.alert(
        `Migration stopped after ${migratedCustomers} customer(s) and ${migratedInvoices} invoice(s). The old data was kept in LocalStorage. Check the records already copied before retrying to avoid duplicates. ${
          error instanceof Error ? error.message : "Please try again."
        }`,
      );
    } finally {
      setIsMigratingOldData(false);
    }
  }

  function updateDocketField(field: keyof DocketForm, value: string) {
    setForm((currentForm) => ({ ...currentForm, [field]: value }));
  }

  function saveDocket() {
    const taxableValue = Number(form.amount);
    const newDocket: Docket = {
      docketNo: form.docketNo.trim(),
      date: form.date,
      destination: form.destination.trim(),
      courier: form.courier,
      mode: form.mode,
      weight: Number(form.weight),
      pieces: Number(form.pieces),
      amount: taxableValue,
    };

    setDockets((currentDockets) => [...currentDockets, newDocket]);
    setForm({ ...emptyDocketForm });
    window.setTimeout(() => {
      document.getElementById("docket-no-input")?.focus();
    }, 0);
  }

  function addDocket(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    saveDocket();
  }

  function handleDocketInputKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.key !== "Enter") return;

    event.preventDefault();
    const currentField =
      event.target instanceof HTMLInputElement ||
      event.target instanceof HTMLSelectElement
        ? event.target
        : null;
    if (!currentField) return;

    const rowFields = currentField
      .closest("tr")
      ?.querySelectorAll<HTMLInputElement | HTMLSelectElement>("input, select");
    const currentIndex = rowFields
      ? Array.from(rowFields).indexOf(currentField)
      : -1;
    const nextField = rowFields?.[currentIndex + 1];

    if (nextField) {
      nextField.focus();
      return;
    }

    if (currentField.form?.reportValidity()) {
      saveDocket();
    }
  }

  function handleDocketSelectKeyDown(
    event: ReactKeyboardEvent<HTMLSelectElement>,
  ) {
    if (event.key !== "Enter") return;

    event.preventDefault();
    if ("showPicker" in HTMLSelectElement.prototype) {
      event.currentTarget.showPicker();
    }
  }

  async function addCustomer(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const newCustomer: Customer = {
      name: customerForm.name.trim(),
      phone: customerForm.phone.trim(),
      address: customerForm.address.trim(),
      state: customerForm.state,
      gstin: customerForm.gstin?.trim() || "",
      contactPerson: customerForm.contactPerson?.trim() || "",
      email: customerForm.email?.trim() || "",
      pan: customerForm.pan?.trim().toUpperCase() || "",
      city: customerForm.city?.trim() || "",
      stateCode: customerForm.stateCode?.trim() || "",
      pincode: customerForm.pincode?.trim() || "",
    };

    try {
      const response = await fetch("/api/customers", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(newCustomer),
      });
      const result: unknown = await response.json();
      if (!response.ok) {
        const message =
          typeof result === "object" && result !== null &&
          typeof (result as Record<string, unknown>).error === "string"
            ? (result as Record<string, string>).error
            : "Customer could not be saved.";
        throw new Error(message);
      }
      if (
        !isCustomer(result) ||
        typeof (result as Record<string, unknown>)._id !== "string"
      ) {
        throw new Error("The server returned invalid customer data.");
      }

      setCustomers((currentCustomers) => [
        {
          ...result,
          gstin: result.gstin ?? "",
          contactPerson: result.contactPerson ?? "",
          email: result.email ?? "",
          pan: result.pan ?? "",
          city: result.city ?? "",
          stateCode: result.stateCode ?? "",
          pincode: result.pincode ?? "",
        },
        ...currentCustomers,
      ]);
      setCustomerForm({ ...emptyCustomerForm });
      setCustomerError("");
      setIsAddingCustomer(false);
    } catch (error) {
      console.error("Could not save customer to MongoDB.", error);
      setCustomerError(
        error instanceof Error ? error.message : "Customer could not be saved. Please try again.",
      );
    }
  }

  async function deleteCustomer(customerToDelete: Customer) {
    if (!customerToDelete._id) {
      setCustomerError("This customer is missing its database ID. Refresh and try again.");
      return;
    }

    try {
      const response = await fetch(`/api/customers/${customerToDelete._id}`, {
        method: "DELETE",
      });
      const result: unknown = await response.json();
      if (!response.ok) {
        const message =
          typeof result === "object" && result !== null &&
          typeof (result as Record<string, unknown>).error === "string"
            ? (result as Record<string, string>).error
            : "Customer could not be deleted.";
        throw new Error(message);
      }

      const updatedCustomers = customers.filter(
        (customer) => customer._id !== customerToDelete._id,
      );
      setCustomers(updatedCustomers);
      if (
        selectedCustomer === customerToDelete.phone &&
        !updatedCustomers.some((customer) => customer.phone === selectedCustomer)
      ) {
        setSelectedCustomer("");
      }
      setCustomerError("");
    } catch (error) {
      console.error("Could not delete customer from MongoDB.", error);
      setCustomerError(
        error instanceof Error ? error.message : "Customer could not be deleted. Please try again.",
      );
    }
  }

  function addCourierPartner(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const newPartner: CourierPartner = {
      id: `${Date.now()}-${courierPartnerForm.name.trim()}`,
      name: courierPartnerForm.name.trim(),
      contactPerson: courierPartnerForm.contactPerson.trim(),
      phone: courierPartnerForm.phone.trim(),
      notes: courierPartnerForm.notes.trim(),
    };
    const updatedPartners = [...courierPartners, newPartner];

    try {
      localStorage.setItem("agc_couriers", JSON.stringify(updatedPartners));
    } catch (error) {
      console.error("Could not save courier partner to localStorage.", error);
      setCourierError("Courier partner could not be saved. Check browser storage and try again.");
      return;
    }

    setCourierPartners(updatedPartners);
    setCourierPartnerForm({ ...emptyCourierPartnerForm });
    setCourierError("");
    setIsAddingCourierPartner(false);
  }

  function deleteCourierPartner(partnerId: string) {
    const updatedPartners = courierPartners.filter((partner) => partner.id !== partnerId);

    try {
      localStorage.setItem("agc_couriers", JSON.stringify(updatedPartners));
    } catch (error) {
      console.error("Could not delete courier partner from localStorage.", error);
      setCourierError("Courier partner could not be deleted. Check browser storage and try again.");
      return;
    }

    setCourierPartners(updatedPartners);
    setCourierError("");
  }

  function updateCompanyField(field: keyof CompanySettings, value: string) {
    setCompanySettings((currentSettings) => ({
      ...currentSettings,
      [field]: value,
    }));
    setSettingsError("");
    setSettingsSuccess("");
  }

  function readCompanyImage(field: "logo" | "signature", file: File | undefined) {
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setSettingsError(`Choose a valid image file for the company ${field}.`);
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") {
        updateCompanyField(field, reader.result);
      } else {
        setSettingsError(`The selected company ${field} could not be read as an image.`);
      }
    };
    reader.onerror = () => {
      console.error(`Could not read the selected company ${field}.`, reader.error);
      setSettingsError(`The selected ${field} could not be read. Please try another image.`);
    };
    reader.readAsDataURL(file);
  }

  function saveCompanySettings(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    try {
      localStorage.setItem("agc_company_settings", JSON.stringify(companySettings));
    } catch (error) {
      console.error("Could not save company settings to localStorage.", error);
      setSettingsError("Settings could not be saved. Check browser storage and try again.");
      setSettingsSuccess("");
      return;
    }

    setSettingsError("");
    setSettingsSuccess("Company profile and bank details saved successfully.");
  }

  function saveAccountSecurity(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const normalizedEmail = loginEmail.trim();
    if (!normalizedEmail) {
      setSecurityError("Enter a login email.");
      setSecuritySuccess("");
      return;
    }
    if (newPassword.length < 8) {
      setSecurityError("New password must be at least 8 characters long.");
      setSecuritySuccess("");
      return;
    }
    if (newPassword !== confirmPassword) {
      setSecurityError("New password and confirmation do not match.");
      setSecuritySuccess("");
      return;
    }

    try {
      localStorage.setItem(
        "agc_auth",
        JSON.stringify({ email: normalizedEmail, password: newPassword } satisfies AuthCredentials),
      );
    } catch (error) {
      console.error("Could not save account credentials to localStorage.", error);
      setSecurityError("Account credentials could not be saved. Check browser storage and try again.");
      setSecuritySuccess("");
      return;
    }

    setLoginEmail(normalizedEmail);
    setNewPassword("");
    setConfirmPassword("");
    setSecurityError("");
    setSecuritySuccess("Account credentials updated successfully.");
  }

  function startNewInvoice() {
    const defaults = getInvoiceDateDefaults();
    setEditingInvoiceId(null);
    setPreviewInvoiceNo(
      createUniqueInvoiceNumber(
        savedInvoices.map((invoice) => invoice.invoiceNo),
        "GST",
      ),
    );
    setDockets([]);
    setForm({ ...emptyDocketForm });
    setInvoiceDate(defaults.invoiceDate);
    setDueDate(defaults.dueDate);
    setInvoiceType("GST");
    setSelectedCustomer("");
    setInvoiceTerms(getDefaultInvoiceTerms(companySettings.city));
    setInvoiceStatus("Issued");
    setPaymentStatus("Unpaid");
    setInvoiceSuccess("");
    setInvoiceError("");
    setIsCreatingInvoice(true);
  }

  function editInvoice(invoice: SavedInvoice) {
    setEditingInvoiceId(invoice._id);
    setActiveTab("invoices");
    setIsCreatingInvoice(true);
    setInvoiceDate(getDateInputValue(invoice.date));
    setDueDate(invoice.dueDate || getInvoiceDateDefaults().dueDate);
    setInvoiceType(invoice.invoiceType ?? "GST");
    setSelectedCustomer(invoice.customer.phone);
    setDockets(invoice.consignments.map((docket) => ({ ...docket })));
    setForm({ ...emptyDocketForm });
    setInvoiceTerms(invoice.terms);
    setInvoiceStatus(invoice.status);
    setPaymentStatus(invoice.paymentStatus);
    setInvoiceSuccess("");
    setInvoiceError("");
  }

  function cancelInvoiceEdit() {
    setIsCreatingInvoice(false);
    setEditingInvoiceId(null);
    setInvoiceError("");
  }

  async function createInvoice() {
    if (!selectedCustomerRecord) {
      setInvoiceError("Select a customer before creating the invoice.");
      return;
    }
    if (dockets.length === 0) {
      setInvoiceError("Add at least one docket before creating the invoice.");
      return;
    }

    const generatedAt = new Date();
    const existingInvoice = editingInvoiceId
      ? savedInvoices.find((savedInvoice) => savedInvoice._id === editingInvoiceId)
      : undefined;
    if (editingInvoiceId && !existingInvoice) {
      setInvoiceError("This invoice could not be found. Return to the invoice list and try again.");
      return;
    }
    let invoiceNo: string;
    try {
      if (existingInvoice && (existingInvoice.invoiceType ?? "GST") === invoiceType) {
        invoiceNo = existingInvoice.invoiceNo;
      } else {
        const existingInvoiceNumbers = savedInvoices
          .filter((savedInvoice) => savedInvoice._id !== existingInvoice?._id)
          .map((savedInvoice) => savedInvoice.invoiceNo);
        invoiceNo = createUniqueInvoiceNumber(existingInvoiceNumbers, invoiceType);
      }
    } catch (error) {
      console.error("Could not generate a unique invoice number.", error);
      setInvoiceError("A unique invoice number could not be generated. Please try again.");
      return;
    }
    const invoiceIsSameState =
      selectedCustomerRecord.state.trim().toLocaleLowerCase() ===
      companySettings.state.trim().toLocaleLowerCase();
    const invoiceGstType = invoiceIsSameState ? "CGST+SGST" : "IGST";
    const invoiceConsignments = dockets.map((docket) => ({
      ...docket,
      ...calculateInclusiveDocketTax(docket.amount, invoiceGstType, invoiceType),
    }));
    const invoiceTaxTotals = calculateInclusiveInvoiceTotals(
      invoiceConsignments,
      invoiceGstType,
      invoiceType,
    );
    const invoiceDateValue = invoiceDate || getInvoiceDateDefaults(generatedAt).invoiceDate;
    const [year, month, day] = invoiceDateValue.split("-").map(Number);
    const formattedDate = new Date(year, month - 1, day).toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
    const invoicePayload = {
      invoiceNo,
      invoiceDate: formattedDate,
      dueDate,
      customerName: selectedCustomerRecord.name,
      customerState: selectedCustomerRecord.state,
      invoiceType,
      gstType: invoiceGstType,
      consignments: invoiceConsignments,
      taxableValue: invoiceTaxTotals.subtotal,
      cgst: invoiceTaxTotals.cgstTotal,
      sgst: invoiceTaxTotals.sgstTotal,
      igst: invoiceTaxTotals.igstTotal,
      totalAmount: invoiceTaxTotals.totalAmount,
      status: invoiceStatus,
      paymentStatus,
      customer: { ...selectedCustomerRecord },
      terms: invoiceTerms,
      gstInclusive: true,
    };

    try {
      const response = await fetch(
        existingInvoice ? `/api/invoices/${existingInvoice._id}` : "/api/invoices",
        {
          method: existingInvoice ? "PUT" : "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(invoicePayload),
        },
      );
      const result: unknown = await response.json();
      if (!response.ok) {
        const message =
          typeof result === "object" && result !== null &&
          typeof (result as Record<string, unknown>).error === "string"
            ? (result as Record<string, string>).error
            : "The invoice could not be saved.";
        throw new Error(message);
      }
      if (!isMongoInvoice(result)) {
        throw new Error("The server returned invalid invoice data.");
      }

      const persistedInvoice = toSavedInvoice(result);
      const updatedInvoices = existingInvoice
        ? savedInvoices.map((savedInvoice) =>
            savedInvoice._id === existingInvoice._id ? persistedInvoice : savedInvoice,
          )
        : [persistedInvoice, ...savedInvoices];
      setSavedInvoices(updatedInvoices);
      setPreviewInvoiceNo(
        createUniqueInvoiceNumber(
          updatedInvoices.map((savedInvoice) => savedInvoice.invoiceNo),
          "GST",
        ),
      );
      setInvoiceSuccess(
        existingInvoice
          ? `Invoice ${invoiceNo} updated successfully.`
          : `Invoice ${invoiceNo} created successfully.`,
      );
      setInvoiceError("");
      setDockets([]);
      setForm({ ...emptyDocketForm });
      const defaults = getInvoiceDateDefaults();
      setInvoiceDate(defaults.invoiceDate);
      setDueDate(defaults.dueDate);
      setInvoiceType("GST");
      setSelectedCustomer("");
      setInvoiceTerms(getDefaultInvoiceTerms(companySettings.city));
      setInvoiceStatus("Issued");
      setPaymentStatus("Unpaid");
      setEditingInvoiceId(null);
      setIsCreatingInvoice(false);
    } catch (error) {
      console.error("Could not save invoice to MongoDB.", error);
      setInvoiceError(
        error instanceof Error ? error.message : "The invoice could not be saved. Please try again.",
      );
    }
  }

  function printSavedInvoice(invoice: SavedInvoice) {
    setInvoiceError("");
    setSelectedPrintInvoice(invoice);
    setAutoPrintInvoice(true);
  }

  function exportFilteredInvoices() {
    try {
      const exportData = filteredInvoices.map((invoice) => ({
        InvoiceNo: invoice.invoiceNo,
        Date: getDateInputValue(invoice.date),
        CustomerName: invoice.customer.name,
        CustomerGSTIN: invoice.customer.gstin ?? "",
        InvoiceType: invoice.invoiceType ?? "GST",
        TaxableValue: roundCurrencyAmount(invoice.subtotal),
        CGST: invoice.invoiceType === "Non-GST" ? 0 : roundCurrencyAmount(invoice.cgstTotal),
        SGST: invoice.invoiceType === "Non-GST" ? 0 : roundCurrencyAmount(invoice.sgstTotal),
        IGST: invoice.invoiceType === "Non-GST" ? 0 : roundCurrencyAmount(invoice.igstTotal),
        TotalAmount: roundCurrencyAmount(invoice.totalAmount),
      }));
      const blob = new Blob([JSON.stringify(exportData, null, 2)], {
        type: "application/json",
      });
      const objectUrl = URL.createObjectURL(blob);
      const downloadLink = document.createElement("a");
      downloadLink.href = objectUrl;
      downloadLink.download = `invoices-${invoiceFilter}-${getInvoiceDateDefaults().invoiceDate}.json`;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      downloadLink.remove();
      window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
      setInvoiceError("");
    } catch (error) {
      console.error("Could not export invoices as JSON.", error);
      setInvoiceError("Invoices could not be exported. Please try again.");
    }
  }

  function exportFilteredInvoicesToExcel() {
    try {
      const exportData = filteredInvoices.map((invoice) => ({
        "Invoice No": invoice.invoiceNo,
        Date: getDateInputValue(invoice.date),
        "Customer Name": invoice.customer.name,
        "Customer GSTIN": invoice.customer.gstin ?? "",
        "Invoice Type": invoice.invoiceType ?? "GST",
        "Taxable Value": roundCurrencyAmount(invoice.subtotal),
        CGST: invoice.invoiceType === "Non-GST" ? 0 : roundCurrencyAmount(invoice.cgstTotal),
        SGST: invoice.invoiceType === "Non-GST" ? 0 : roundCurrencyAmount(invoice.sgstTotal),
        IGST: invoice.invoiceType === "Non-GST" ? 0 : roundCurrencyAmount(invoice.igstTotal),
        "Total Amount": roundCurrencyAmount(invoice.totalAmount),
        Status: invoice.status,
      }));
      const worksheet = XLSX.utils.json_to_sheet(exportData);
      worksheet["!cols"] = [
        { wch: 16 },
        { wch: 14 },
        { wch: 24 },
        { wch: 20 },
        { wch: 14 },
        { wch: 16 },
        { wch: 12 },
        { wch: 12 },
        { wch: 12 },
        { wch: 16 },
        { wch: 14 },
      ];
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Invoices");
      const filterName =
        invoiceFilter === "all" ? "All" : invoiceFilter === "GST" ? "GST" : "Non-GST";
      XLSX.writeFile(workbook, `${filterName}_Invoices.xlsx`);
      setInvoiceError("");
    } catch (error) {
      console.error("Could not export invoices as an Excel workbook.", error);
      setInvoiceError("Invoices could not be exported to Excel. Please try again.");
    }
  }

  async function deleteInvoice(invoiceId: string) {
    if (!window.confirm("Are you sure you want to delete this invoice?")) {
      return;
    }

    try {
      const response = await fetch(`/api/invoices/${invoiceId}`, { method: "DELETE" });
      const result: unknown = await response.json();
      if (!response.ok) {
        const message =
          typeof result === "object" && result !== null &&
          typeof (result as Record<string, unknown>).error === "string"
            ? (result as Record<string, string>).error
            : "The invoice could not be deleted.";
        throw new Error(message);
      }

      setSavedInvoices((currentInvoices) =>
        currentInvoices.filter((invoice) => invoice._id !== invoiceId),
      );
      setInvoiceError("");
    } catch (error) {
      console.error("Could not delete invoice from MongoDB.", error);
      setInvoiceError(
        error instanceof Error ? error.message : "The invoice could not be deleted. Please try again.",
      );
    }
  }

  async function addTeamUser(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setTeamError("");

    try {
      const response = await fetch("/api/auth/register", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(teamForm),
      });
      const result: unknown = await response.json();
      if (!response.ok) {
        const message =
          typeof result === "object" &&
          result !== null &&
          typeof (result as Record<string, unknown>).error === "string"
            ? (result as Record<string, string>).error
            : "Could not create user.";
        throw new Error(message);
      }

      const user =
        typeof result === "object" && result !== null
          ? (result as Record<string, unknown>).user
          : null;
      if (!isTeamUser(user)) {
        throw new Error("The registration API returned invalid user data.");
      }

      setTeamUsers((currentUsers) => [...currentUsers, user]);
      setTeamForm({ name: "", email: "", password: "", role: "sub-user" });
      setIsAddingTeamUser(false);
    } catch (error) {
      console.error("Could not create a team user.", error);
      setTeamError(error instanceof Error ? error.message : "Could not create user.");
    }
  }

  async function login(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setIsLoggingIn(true);
    setLoginError("");

    try {
      const response = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: username.trim(), password }),
      });
      const result: unknown = await response.json();

      if (!response.ok) {
        const message =
          typeof result === "object" &&
          result !== null &&
          typeof (result as Record<string, unknown>).error === "string"
            ? (result as Record<string, string>).error
            : "Login could not be verified. Please try again.";
        setLoginError(message);
        return;
      }

      const user =
        typeof result === "object" && result !== null
          ? (result as Record<string, unknown>).user
          : null;
      if (!isLoggedInUser(user)) {
        throw new Error("The login service returned invalid user data.");
      }

      localStorage.setItem("agc_user", JSON.stringify(user));
      setLoggedInUser(user);
      setLoginEmail(user.email);
      setIsLoggedIn(true);
      setPassword("");
    } catch (error) {
      console.error("Could not complete database-backed login.", error);
      setLoginError(
        error instanceof Error ? error.message : "Login could not be verified. Please try again.",
      );
    } finally {
      setIsLoggingIn(false);
    }
  }

  function logout() {
    void fetch("/api/auth/logout", { method: "POST" }).catch((error: unknown) => {
      console.error("Could not end the server session.", error);
    });
    localStorage.removeItem("agc_user");
    setLoggedInUser(null);
    setIsLoggedIn(false);
  }

  function openInvoices() {
    setActiveTab("invoices");
    setIsCreatingInvoice(false);
    setInvoiceSuccess("");
    setInvoiceError("");
  }

  function generateReport() {
    if (!reportStartDate || !reportEndDate) {
      setReportError("Select both a start date and an end date.");
      return;
    }
    if (reportStartDate > reportEndDate) {
      setReportError("Start date must be on or before the end date.");
      return;
    }

    setGeneratedReportRange({ startDate: reportStartDate, endDate: reportEndDate });
    setReportError("");
  }

  return (
    <main className="min-h-screen bg-slate-50 text-slate-800 print:min-h-0 print:bg-white print:p-0 print:text-black">
      {!isLoggedIn ? (
        <section className="grid min-h-screen w-full bg-white md:grid-cols-2 print:hidden">
          <div className="relative hidden min-h-screen overflow-hidden bg-gradient-to-br from-blue-900 to-slate-900 px-12 py-16 text-white md:flex md:flex-col md:justify-between lg:px-20">
            <div
              aria-hidden="true"
              className="absolute -left-36 -top-32 h-[28rem] w-[28rem] rounded-full bg-blue-500/20 blur-3xl"
            />
            <div
              aria-hidden="true"
              className="absolute -bottom-48 -right-32 h-[34rem] w-[34rem] rounded-full bg-cyan-400/10 blur-3xl"
            />
            <div
              aria-hidden="true"
              className="absolute right-16 top-1/3 h-56 w-56 rounded-full border border-white/10"
            />
            <div className="relative z-10 flex items-center gap-4">
              <div className="flex h-16 w-28 shrink-0 items-center justify-center overflow-hidden rounded-2xl border border-white/20 bg-white p-2 shadow-lg">
                {companySettings.logo ? (
                  // eslint-disable-next-line @next/next/no-img-element
                  <img
                    src={companySettings.logo}
                    alt={`${companySettings.name || "SD Enterprises"} logo`}
                    className="h-full w-full object-contain"
                  />
                ) : (
                  <span className="text-2xl font-bold tracking-tight text-blue-900">
                    SD
                  </span>
                )}
              </div>
              <div>
                <p className="text-lg font-bold tracking-[0.18em]">SD</p>
                <p className="text-xs font-semibold tracking-[0.28em] text-blue-100/80">
                  ENTERPRISES
                </p>
              </div>
            </div>
            <div className="relative z-10 max-w-xl py-16">
              <p className="mb-5 text-sm font-semibold uppercase tracking-[0.2em] text-blue-200">
                Courier billing, simplified
              </p>
              <h2 className="text-4xl font-semibold leading-tight tracking-tight lg:text-5xl">
                Streamline your courier billing.
              </h2>
              <p className="mt-6 max-w-lg text-lg leading-8 text-slate-300">
                Manage invoices, customers, and GST filings in one powerful
                dashboard.
              </p>
            </div>
            <p className="relative z-10 text-sm text-slate-400">
              A clearer view of your business, from docket to payment.
            </p>
          </div>

          <div className="flex min-h-screen items-center justify-center bg-slate-50 px-5 py-12 sm:px-8">
            <div className="w-full max-w-md rounded-2xl bg-white p-7 shadow-xl shadow-slate-200/70 ring-1 ring-slate-200/70 sm:p-10">
              <div className="mb-8">
                <div className="mb-8 flex items-center gap-3 md:hidden">
                  <div className="flex h-12 w-20 shrink-0 items-center justify-center overflow-hidden rounded-xl border border-slate-200 bg-white p-1.5 shadow-sm">
                    {companySettings.logo ? (
                      // eslint-disable-next-line @next/next/no-img-element
                      <img
                        src={companySettings.logo}
                        alt={`${companySettings.name || "SD Enterprises"} logo`}
                        className="h-full w-full object-contain"
                      />
                    ) : (
                      <span className="text-lg font-bold text-blue-900">SD</span>
                    )}
                  </div>
                  <div>
                    <p className="text-sm font-extrabold tracking-[0.12em] text-slate-900">
                      SD ENTERPRISES
                    </p>
                    <p className="text-[10px] font-semibold tracking-[0.16em] text-slate-500">
                      BILLING SUITE
                    </p>
                  </div>
                </div>
                <h1 className="text-2xl font-bold tracking-tight text-slate-900">
                  Welcome back
                </h1>
                <p className="mt-2 text-sm text-gray-500">
                  Sign in to your account
                </p>
              </div>

              <form onSubmit={login} className="space-y-5">
                <label className="block text-sm font-medium text-slate-700">
                  Email address
                  <span className="relative mt-2 block">
                    <svg
                      aria-hidden="true"
                      className="pointer-events-none absolute left-3 top-1/2 h-5 w-5 -translate-y-1/2 text-slate-400"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="1.8"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    >
                      <rect x="3" y="5" width="18" height="14" rx="2" />
                      <path d="m3 7 9 6 9-6" />
                    </svg>
                    <input
                      required
                      type="email"
                      autoComplete="username"
                      value={username}
                      onChange={(event) => setUsername(event.target.value)}
                      disabled={isLoggingIn}
                      className="w-full rounded-lg border border-gray-200 bg-gray-50 py-3 pl-10 pr-3 text-sm text-slate-900 outline-none transition placeholder:text-slate-400 focus:border-blue-600 focus:bg-white focus:ring-2 focus:ring-blue-600/20 disabled:cursor-not-allowed disabled:opacity-60"
                    />
                  </span>
                </label>

                <label className="block text-sm font-medium text-slate-700">
                  Password
                  <span className="relative mt-2 block">
                    <svg
                      aria-hidden="true"
                      className="pointer-events-none absolute left-3 top-1/2 h-5 w-5 -translate-y-1/2 text-slate-400"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="1.8"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    >
                      <rect x="4" y="10" width="16" height="11" rx="2" />
                      <path d="M8 10V7a4 4 0 0 1 8 0v3" />
                      <path d="M12 14v3" />
                    </svg>
                    <input
                      required
                      type="password"
                      autoComplete="current-password"
                      value={password}
                      onChange={(event) => setPassword(event.target.value)}
                      disabled={isLoggingIn}
                      className="w-full rounded-lg border border-gray-200 bg-gray-50 py-3 pl-10 pr-3 text-sm text-slate-900 outline-none transition placeholder:text-slate-400 focus:border-blue-600 focus:bg-white focus:ring-2 focus:ring-blue-600/20 disabled:cursor-not-allowed disabled:opacity-60"
                    />
                  </span>
                </label>

                <div className="flex items-center justify-between gap-4 text-sm">
                  <label className="flex items-center gap-2 text-slate-600">
                    <input
                      type="checkbox"
                      className="h-4 w-4 rounded border-gray-300 text-blue-700 accent-blue-700 focus:ring-blue-600"
                    />
                    Remember me
                  </label>
                  <a
                    href="#"
                    onClick={(event) => event.preventDefault()}
                    className="font-medium text-blue-700 transition hover:text-blue-800"
                  >
                    Forgot password?
                  </a>
                </div>

                {loginError && (
                  <p role="alert" className="text-sm font-medium text-red-600">
                    {loginError}
                  </p>
                )}
                <button
                  type="submit"
                  disabled={isLoggingIn}
                  className="w-full rounded-lg bg-blue-700 px-5 py-3 text-sm font-semibold text-white shadow-md shadow-blue-900/15 transition hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-200 disabled:cursor-not-allowed disabled:opacity-70"
                >
                  {isLoggingIn ? "Signing in..." : "Sign in"}
                </button>
              </form>
              <p className="mt-8 text-center text-xs text-slate-400">
                SD Enterprises Billing Suite
              </p>
            </div>
          </div>
        </section>
      ) : (
        <div className="app-theme min-h-screen print:hidden">
          <aside className="fixed inset-y-0 left-0 z-30 flex h-screen w-64 flex-col bg-slate-900 px-4 py-6 text-white print:hidden">
            <div className="flex items-center gap-3 border-b border-slate-800 px-2 pb-6">
              <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-600 text-white shadow-lg shadow-blue-950/30">
                <svg
                  aria-hidden="true"
                  className="h-5 w-5"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M3 7h11v11H3zM14 11h4l3 3v4h-7z" />
                  <circle cx="7.5" cy="18" r="2" />
                  <circle cx="17.5" cy="18" r="2" />
                </svg>
              </div>
              <div className="min-w-0">
                <p className="text-[11px] font-extrabold leading-4 tracking-[0.09em]">
                  SD ENTERPRISES
                </p>
                <p className="text-[10px] font-semibold leading-4 tracking-[0.12em] text-slate-400">
                  BILLING SUITE
                </p>
              </div>
            </div>

            <p className="mt-7 px-3 text-[10px] font-bold uppercase tracking-[0.16em] text-slate-500">
              Workspace
            </p>
            <nav className="mt-3 flex-1 space-y-1" aria-label="Main navigation">
              {navigationItems
                .filter((item) => item.id !== "team" || loggedInUser?.role === "admin")
                .map((item) => {
                  const isActive = activeTab === item.id;

                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => {
                        setActiveTab(item.id);
                        if (item.id === "team") {
                          setIsAddingTeamUser(false);
                          setTeamError("");
                        }
                        if (item.id === "invoices") {
                          setIsCreatingInvoice(false);
                          setInvoiceSuccess("");
                          setInvoiceError("");
                        }
                        if (item.id === "customers") {
                          setIsAddingCustomer(false);
                          setCustomerError("");
                        }
                      }}
                      aria-current={isActive ? "page" : undefined}
                      className={`flex w-full items-center gap-3 rounded-lg border-l-2 px-3 py-2.5 text-left text-sm font-medium transition ${
                        isActive
                          ? "border-blue-400 bg-blue-600 text-white shadow-sm"
                          : "border-transparent text-slate-400 hover:bg-slate-800 hover:text-white"
                      }`}
                    >
                      <NavIcon name={item.id} />
                      <span>{item.label}</span>
                    </button>
                  );
                })}
            </nav>

            <div className="border-t border-slate-800 pt-4">
              <button
                type="button"
                onClick={logout}
                className="flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-red-400 transition hover:bg-red-500/10 hover:text-red-300"
              >
                <svg
                  aria-hidden="true"
                  className="h-5 w-5"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.8"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M10 17l5-5-5-5M15 12H3M12 3h6a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-6" />
                </svg>
                Logout
              </button>
              <p className="mt-4 truncate px-3 text-[11px] text-slate-500">
                {loggedInUser?.email}
              </p>
            </div>
          </aside>

          <div className="ml-64 min-h-screen">
            <header className="flex h-[68px] items-center justify-end border-b border-slate-200 bg-white px-6 sm:px-8">
              <div className="flex items-center gap-3">
                <span className="text-right text-sm text-slate-600">
                  <span className="block font-medium text-slate-800">
                    {loggedInUser?.name}
                  </span>
                  <span className="block text-xs capitalize text-slate-500">
                    {loggedInUser?.role === "admin" ? "Admin" : "Sub-user"}
                  </span>
                </span>
                <span className="flex h-9 w-9 items-center justify-center rounded-full bg-blue-600 text-sm font-semibold uppercase text-white ring-4 ring-blue-50">
                  {loggedInUser?.name.charAt(0) ?? "U"}
                </span>
              </div>
            </header>

            <div className="mx-auto max-w-[1500px] px-6 py-8 sm:px-8 lg:px-10">
              {activeTab === "dashboard" && (
                <>
                  <div className="mb-7 flex flex-wrap items-end justify-between gap-4">
                    <div>
                      <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500">
                        OVERVIEW
                      </p>
                      <h1 className="mt-1 text-3xl font-bold tracking-tight text-slate-950">
                        Dashboard
                      </h1>
                    </div>
                    <div className="flex flex-wrap gap-3">
                      <button
                        type="button"
                        onClick={() => void migrateOldDataToDB()}
                        disabled={isMigratingOldData}
                        className="inline-flex items-center gap-2 rounded-md border border-amber-300 bg-amber-50 px-4 py-2.5 text-sm font-semibold text-amber-800 shadow-sm transition hover:bg-amber-100 disabled:cursor-not-allowed disabled:opacity-60"
                      >
                        {isMigratingOldData ? "Migrating..." : "Migrate Old Data"}
                      </button>
                      <button
                        type="button"
                        onClick={() => setActiveTab("settings")}
                        className="inline-flex items-center gap-2 rounded-md border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50"
                      >
                        <svg
                          aria-hidden="true"
                          className="h-4 w-4"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.8"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <circle cx="12" cy="12" r="3" />
                          <path d="m19.4 15 .1.1 1.4 1.1-1.4 2.4-1.7-.6a7.8 7.8 0 0 1-1.5.9l-.3 1.8h-2.8l-.3-1.8a7.8 7.8 0 0 1-1.5-.9l-1.7.6-1.4-2.4 1.4-1.1a7 7 0 0 1 0-1.8l-1.4-1.1 1.4-2.4 1.7.6a7.8 7.8 0 0 1 1.5-.9l.3-1.8h2.8l.3 1.8a7.8 7.8 0 0 1 1.5.9l1.7-.6 1.4 2.4-1.4 1.1a7 7 0 0 1 0 1.8Z" />
                        </svg>
                        My Company
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setActiveTab("invoices");
                          startNewInvoice();
                        }}
                        className="inline-flex items-center gap-2 rounded-md bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-100"
                      >
                        <span className="text-lg leading-none">+</span>
                        New Invoice
                      </button>
                    </div>
                  </div>

                  <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
                    {[
                      {
                        label: "INVOICES",
                        value: String(savedInvoices.length),
                        detail: "Invoices generated",
                        color: "border-blue-500",
                        icon: "▤",
                      },
                      {
                        label: "BILLED",
                        value: formatCurrency(
                          savedInvoices.reduce((total, invoice) => total + invoice.totalAmount, 0),
                        ),
                        detail: "Including GST",
                        color: "border-emerald-500",
                        icon: "₹",
                      },
                      {
                        label: "COLLECTED",
                        value: formatCurrency(
                          savedInvoices
                            .filter((invoice) => invoice.paymentStatus === "Paid")
                            .reduce((total, invoice) => total + invoice.totalAmount, 0),
                        ),
                        detail: "Payments received",
                        color: "border-orange-400",
                        icon: "↗",
                      },
                      {
                        label: "OUTSTANDING",
                        value: formatCurrency(
                          savedInvoices
                            .filter((invoice) => invoice.paymentStatus !== "Paid")
                            .reduce((total, invoice) => total + invoice.totalAmount, 0),
                        ),
                        detail: "Awaiting payment",
                        color: "border-violet-500",
                        icon: "◷",
                      },
                    ].map((stat) => (
                      <article
                        key={stat.label}
                        className={`rounded-lg border border-slate-200 border-b-4 ${stat.color} bg-white p-5 shadow-sm`}
                      >
                        <div className="flex items-start justify-between">
                          <p className="text-[11px] font-bold tracking-[0.12em] text-slate-500">
                            {stat.label}
                          </p>
                          <span className="flex h-9 w-9 items-center justify-center rounded-lg bg-slate-50 text-lg font-semibold text-slate-600">
                            {stat.icon}
                          </span>
                        </div>
                        <p className="mt-4 truncate text-2xl font-bold tracking-tight text-slate-950">
                          {stat.value}
                        </p>
                        <p className="mt-1 text-xs text-slate-500">{stat.detail}</p>
                      </article>
                    ))}
                  </div>

                  <section className="mt-7 overflow-hidden rounded-2xl border border-slate-200/80 bg-gradient-to-br from-white via-white to-blue-50/40 p-5 shadow-[0_8px_30px_rgb(15,23,42,0.04)] transition-shadow duration-300 hover:shadow-[0_12px_36px_rgb(15,23,42,0.08)] sm:p-6">
                    <div className="mb-5 flex flex-wrap items-center justify-between gap-4">
                      <div className="flex items-center gap-3">
                        <span className="flex h-11 w-11 items-center justify-center rounded-xl bg-gradient-to-br from-blue-500 to-indigo-600 text-white shadow-md shadow-blue-500/20">
                          <svg
                            aria-hidden="true"
                            className="h-5 w-5"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="1.8"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <path d="M3 3v18h18" />
                            <path d="m19 9-5 5-4-4-5 5" />
                          </svg>
                        </span>
                        <div>
                          <h2 className="text-lg font-bold tracking-tight text-slate-900">
                            Monthly Revenue
                          </h2>
                          <p className="mt-1 text-sm text-slate-500">Total sales overview</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-3">
                        <span className="rounded-full border border-blue-100 bg-blue-50/80 px-3 py-1.5 text-xs font-semibold text-blue-700">
                          Last 6 months
                        </span>
                        <span className="hidden text-right sm:block">
                          <span className="block text-[10px] font-semibold uppercase tracking-wider text-slate-400">
                            Revenue
                          </span>
                          <span className="mt-0.5 block text-sm font-bold tabular-nums text-slate-900">
                            {formatCurrency(
                              monthlyRevenueData.reduce((sum, month) => sum + month.revenue, 0),
                            )}
                          </span>
                        </span>
                      </div>
                    </div>
                    {monthlyRevenueData.length > 0 ? (
                      <div className="w-full">
                        <ResponsiveContainer width="100%" height={350}>
                          <AreaChart
                            data={monthlyRevenueData}
                            margin={{ top: 12, right: 12, left: 12, bottom: 4 }}
                          >
                            <defs>
                              <linearGradient id="monthlyRevenueGradient" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="0%" stopColor="#3b82f6" stopOpacity={0.38} />
                                <stop offset="90%" stopColor="#3b82f6" stopOpacity={0.04} />
                                <stop offset="100%" stopColor="#3b82f6" stopOpacity={0} />
                              </linearGradient>
                            </defs>
                            <CartesianGrid
                              stroke="#e2e8f0"
                              strokeDasharray="3 3"
                              vertical={false}
                            />
                            <XAxis
                              dataKey="month"
                              axisLine={false}
                              tickLine={false}
                              tick={{ fill: "#64748b", fontSize: 12 }}
                              tickMargin={10}
                            />
                            <YAxis
                              axisLine={false}
                              tickLine={false}
                              tick={{ fill: "#64748b", fontSize: 12 }}
                              tickFormatter={(value: number) => `₹${currencyFormatter.format(value)}`}
                              width={88}
                            />
                            <Tooltip
                              cursor={{ stroke: "#94a3b8", strokeDasharray: "4 4" }}
                              content={({ active, payload, label }) => {
                                if (!active || !payload?.length) return null;
                                return (
                                  <div className="rounded-xl border border-slate-700 bg-slate-800 px-4 py-3 text-white shadow-xl">
                                    <p className="text-xs font-medium text-slate-300">{label}</p>
                                    <p className="mt-1 text-base font-bold">
                                      {formatCurrency(Number(payload[0].value))}
                                    </p>
                                  </div>
                                );
                              }}
                            />
                            <Area
                              type="monotone"
                              dataKey="revenue"
                              name="Revenue"
                              fill="url(#monthlyRevenueGradient)"
                              stroke="#3b82f6"
                              strokeWidth={3}
                              dot={false}
                              activeDot={{
                                r: 7,
                                fill: "#3b82f6",
                                stroke: "#ffffff",
                                strokeWidth: 3,
                                style: { filter: "drop-shadow(0 2px 6px rgba(37, 99, 235, 0.35))" },
                              }}
                            />
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>
                    ) : (
                      <div className="flex h-64 items-center justify-center rounded-md bg-slate-50 text-sm text-slate-500">
                        Monthly revenue will appear here after you create invoices.
                      </div>
                    )}
                  </section>

                  <div className="mt-7 grid gap-5 xl:grid-cols-[1.6fr_1fr]">
                    <section className="rounded-lg border border-slate-200 bg-white shadow-sm">
                      <div className="flex items-center justify-between border-b border-slate-100 px-5 py-4">
                        <div>
                          <h2 className="font-semibold text-slate-900">Recent invoices</h2>
                          <p className="mt-1 text-xs text-slate-500">
                            Your latest billing activity
                          </p>
                        </div>
                        <button
                          type="button"
                          onClick={() => {
                            setActiveTab("invoices");
                            startNewInvoice();
                          }}
                          className="text-sm font-semibold text-blue-600 hover:text-blue-700"
                        >
                          View all
                        </button>
                      </div>
                      {savedInvoices.length > 0 ? (
                        <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-5">
                          <div>
                            <p className="font-semibold text-slate-900">
                              {savedInvoices[0].invoiceNo}
                            </p>
                            <p className="mt-1 text-xs text-slate-500">
                              {savedInvoices[0].date} · {savedInvoices[0].customer.name}
                            </p>
                          </div>
                          <div className="text-right">
                            <p className="font-semibold text-slate-900">
                              {formatCurrency(savedInvoices[0].totalAmount)}
                            </p>
                            <span className="mt-1 inline-flex rounded-full bg-emerald-50 px-2.5 py-1 text-xs font-medium text-emerald-700">
                              {savedInvoices[0].status}
                            </span>
                          </div>
                        </div>
                      ) : (
                        <div className="px-5 py-12 text-center">
                          <div className="mx-auto flex h-11 w-11 items-center justify-center rounded-full bg-slate-100 text-slate-500">
                            <NavIcon name="invoices" />
                          </div>
                          <p className="mt-3 text-sm text-slate-500">
                            No invoices yet — create your first one.
                          </p>
                          <button
                            type="button"
                            onClick={() => {
                              setActiveTab("invoices");
                              startNewInvoice();
                            }}
                            className="mt-4 text-sm font-semibold text-blue-600 hover:text-blue-700"
                          >
                            Create invoice
                          </button>
                        </div>
                      )}
                    </section>

                    <section className="rounded-lg border border-slate-200 bg-white shadow-sm">
                      <div className="border-b border-slate-100 px-5 py-4">
                        <h2 className="font-semibold text-slate-900">Masters</h2>
                        <p className="mt-1 text-xs text-slate-500">
                          Manage your business records
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          setActiveTab("customers");
                          setIsAddingCustomer(false);
                          setCustomerError("");
                        }}
                        className="flex w-full items-center justify-between border-b border-slate-100 px-5 py-4 text-left transition hover:bg-slate-50"
                      >
                        <span className="text-sm font-medium text-slate-700">Customers</span>
                        <span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs font-semibold text-slate-600">
                          {customers.length}
                        </span>
                      </button>
                      <button
                        type="button"
                        onClick={() => setActiveTab("couriers")}
                        className="flex w-full items-center justify-between px-5 py-4 text-left transition hover:bg-slate-50"
                      >
                        <span className="text-sm font-medium text-slate-700">
                          Courier Partners
                        </span>
                        <span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs font-semibold text-slate-600">
                          3
                        </span>
                      </button>
                      <div className="border-t border-slate-100 p-4">
                        <button
                          type="button"
                          onClick={() => setActiveTab("reports")}
                          className="w-full rounded-md border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 transition hover:bg-slate-50"
                        >
                          Open Reports
                        </button>
                      </div>
                    </section>
                  </div>
                </>
              )}

              {activeTab === "team" && loggedInUser?.role === "admin" && (
                <section className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
                  <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 px-5 py-5 sm:px-6">
                    <div>
                      <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500">
                        ADMINISTRATION
                      </p>
                      <h1 className="mt-1 text-2xl font-bold tracking-tight text-slate-950">
                        Team
                      </h1>
                      <p className="mt-1 text-sm text-slate-500">
                        Manage users who can access the billing workspace.
                      </p>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        setTeamForm({ name: "", email: "", password: "", role: "sub-user" });
                        setTeamError("");
                        setIsAddingTeamUser(true);
                      }}
                      className="inline-flex items-center gap-2 rounded-md bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700"
                    >
                      <span className="text-lg leading-none">+</span>
                      Add User
                    </button>
                  </div>

                  {teamError && !isAddingTeamUser && (
                    <p role="alert" className="mx-5 mt-4 rounded-md bg-red-50 p-3 text-sm text-red-700 sm:mx-6">
                      {teamError}
                    </p>
                  )}

                  <div className="overflow-x-auto">
                    <table className="w-full min-w-[700px] border-collapse text-left text-sm">
                      <thead className="bg-slate-50 text-slate-600">
                        <tr>
                          {["Name", "Email", "Role"].map((heading) => (
                            <th
                              key={heading}
                              className="border-b border-slate-200 px-5 py-3 text-xs font-semibold uppercase tracking-wide"
                            >
                              {heading}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {isLoadingTeamUsers ? (
                          <tr>
                            <td colSpan={3} className="px-5 py-14 text-center text-slate-500">
                              <span role="status" className="inline-flex items-center gap-2">
                                <span className="h-4 w-4 animate-spin rounded-full border-2 border-slate-300 border-t-blue-600" />
                                Loading users...
                              </span>
                            </td>
                          </tr>
                        ) : teamUsers.length === 0 ? (
                          <tr>
                            <td colSpan={3} className="px-5 py-14 text-center text-slate-500">
                              No users found.
                            </td>
                          </tr>
                        ) : (
                          teamUsers.map((user) => (
                            <tr key={user._id}>
                              <td className="px-5 py-4 font-semibold text-slate-900">
                                {user.name}
                              </td>
                              <td className="px-5 py-4 text-slate-600">{user.email}</td>
                              <td className="px-5 py-4">
                                <span className="inline-flex rounded-full bg-blue-50 px-2.5 py-1 text-xs font-medium capitalize text-blue-700">
                                  {user.role === "admin" ? "Admin" : "Sub-user"}
                                </span>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>

                  {isAddingTeamUser &&
                    createPortal(
                      <div
                        className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
                        onMouseDown={(event) => {
                          if (event.target === event.currentTarget) setIsAddingTeamUser(false);
                        }}
                      >
                        <section
                          role="dialog"
                          aria-modal="true"
                          aria-labelledby="add-team-user-title"
                          className="w-full max-w-md rounded-xl bg-white p-6 shadow-2xl"
                        >
                          <div className="flex items-start justify-between gap-4">
                            <div>
                              <h2 id="add-team-user-title" className="text-xl font-bold text-slate-900">
                                Add User
                              </h2>
                              <p className="mt-1 text-sm text-slate-500">
                                Create login credentials for a team member.
                              </p>
                            </div>
                            <button
                              type="button"
                              aria-label="Close dialog"
                              onClick={() => setIsAddingTeamUser(false)}
                              className="rounded-md p-1 text-xl leading-none text-slate-500 hover:bg-slate-100 hover:text-slate-800"
                            >
                              ×
                            </button>
                          </div>
                          <form onSubmit={addTeamUser} className="mt-5 space-y-4">
                            <label className="block text-sm font-medium text-slate-700">
                              Name
                              <input
                                required
                                autoComplete="name"
                                value={teamForm.name}
                                onChange={(event) =>
                                  setTeamForm((formState) => ({
                                    ...formState,
                                    name: event.target.value,
                                  }))
                                }
                                className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                              />
                            </label>
                            <label className="block text-sm font-medium text-slate-700">
                              Email
                              <input
                                required
                                type="email"
                                autoComplete="email"
                                value={teamForm.email}
                                onChange={(event) =>
                                  setTeamForm((formState) => ({
                                    ...formState,
                                    email: event.target.value,
                                  }))
                                }
                                className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                              />
                            </label>
                            <label className="block text-sm font-medium text-slate-700">
                              Password
                              <input
                                required
                                type="password"
                                minLength={8}
                                autoComplete="new-password"
                                value={teamForm.password}
                                onChange={(event) =>
                                  setTeamForm((formState) => ({
                                    ...formState,
                                    password: event.target.value,
                                  }))
                                }
                                className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                              />
                            </label>
                            <label className="block text-sm font-medium text-slate-700">
                              Role
                              <select
                                value={teamForm.role}
                                onChange={(event) =>
                                  setTeamForm((formState) => ({
                                    ...formState,
                                    role: event.target.value as LoggedInUser["role"],
                                  }))
                                }
                                className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                              >
                                <option value="sub-user">Sub-user</option>
                                <option value="admin">Admin</option>
                              </select>
                            </label>
                            {teamError && (
                              <p role="alert" className="rounded-md bg-red-50 p-3 text-sm text-red-700">
                                {teamError}
                              </p>
                            )}
                            <div className="flex justify-end gap-3 pt-2">
                              <button
                                type="button"
                                onClick={() => setIsAddingTeamUser(false)}
                                className="rounded-md border border-slate-300 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-50"
                              >
                                Cancel
                              </button>
                              <button
                                type="submit"
                                className="rounded-md bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-700"
                              >
                                Save User
                              </button>
                            </div>
                          </form>
                        </section>
                      </div>,
                      document.body,
                    )}
                </section>
              )}

              {activeTab === "invoices" && (
                isCreatingInvoice ? (
                <section>
                  <div className="mb-6 flex flex-wrap items-center justify-between gap-4">
                    <div className="flex items-center gap-4">
                      <button
                        type="button"
                        onClick={cancelInvoiceEdit}
                        className="rounded-md border border-gray-300 bg-white px-3 py-2 text-sm font-medium text-gray-700 shadow-sm transition hover:bg-gray-50"
                      >
                        &lt; Back
                      </button>
                      <h1 className="text-3xl font-bold tracking-tight text-gray-950">
                        {editingInvoiceId ? "Edit Invoice" : "New Invoice"}
                      </h1>
                    </div>
                    <p className="text-sm font-semibold tracking-wide text-gray-500">
                      {editingInvoiceId ? "INVOICE #:" : "NEXT #:"}{" "}
                      <span className="text-gray-900">
                        {editingInvoiceId
                          ? savedInvoices.find((invoice) => invoice._id === editingInvoiceId)
                              ?.invoiceNo ?? previewInvoiceNo
                          : previewInvoiceNo}
                      </span>
                    </p>
                  </div>

                  <div className="space-y-6">
                    <section className="w-full overflow-hidden rounded-xl border border-gray-200 bg-white shadow-sm">
                      <div className="border-b border-gray-200 px-5 py-5 sm:px-6">
                        <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-4">
                          <label className="block text-sm font-medium text-gray-700">
                            Invoice Date
                            <input
                              type="date"
                              value={invoiceDate}
                              onChange={(event) => setInvoiceDate(event.target.value)}
                              className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm text-gray-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            />
                          </label>
                          <label className="block text-sm font-medium text-gray-700">
                            Due Date
                            <input
                              type="date"
                              value={dueDate}
                              onChange={(event) => setDueDate(event.target.value)}
                              className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm text-gray-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            />
                          </label>
                          <label className="block text-sm font-medium text-gray-700">
                            Customer <span className="text-red-500">*</span>
                            <select
                              required
                              value={selectedCustomer}
                              onChange={(event) => setSelectedCustomer(event.target.value)}
                              className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm text-gray-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            >
                              <option value="" disabled>
                                Select a customer
                              </option>
                              {customers.map((customer, index) => (
                                <option key={`${customer.phone}-${index}`} value={customer.phone}>
                                  {customer.name}
                                </option>
                              ))}
                              {editingInvoiceId &&
                                selectedCustomerRecord &&
                                !customers.some(
                                  (customer) => customer.phone === selectedCustomerRecord.phone,
                                ) && (
                                  <option value={selectedCustomerRecord.phone}>
                                    {selectedCustomerRecord.name} (saved on invoice)
                                  </option>
                                )}
                            </select>
                          </label>
                          <label className="block text-sm font-medium text-gray-700">
                            Invoice Type
                            <select
                              value={invoiceType}
                              onChange={(event) => {
                                const nextInvoiceType = event.target.value as InvoiceType;
                                setInvoiceType(nextInvoiceType);
                                if (!editingInvoiceId) {
                                  setPreviewInvoiceNo(
                                    createUniqueInvoiceNumber(
                                      savedInvoices.map((invoice) => invoice.invoiceNo),
                                      nextInvoiceType,
                                    ),
                                  );
                                }
                              }}
                              className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm text-gray-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            >
                              <option value="GST">GST</option>
                              <option value="Non-GST">Non-GST</option>
                            </select>
                          </label>
                        </div>
                      </div>

                      <form onSubmit={addDocket} className="p-5 sm:p-6">
                        <div className="mb-4 flex items-center justify-between gap-3">
                          <h2 className="text-xs font-bold tracking-[0.14em] text-gray-500">
                            CONSIGNMENTS
                          </h2>
                          <button
                            type="submit"
                            className="inline-flex items-center gap-1.5 rounded-md border border-gray-300 bg-white px-3 py-2 text-sm font-semibold text-gray-700 shadow-sm transition hover:bg-gray-50"
                          >
                            <span className="text-lg leading-none">+</span>
                            Add row
                          </button>
                        </div>

                        <div className="rounded-md border border-gray-200">
                          <table className="w-full table-fixed border-collapse text-left text-sm">
                            <thead className="bg-gray-50 text-gray-500">
                              <tr>
                                {[
                                  "DOCKET #",
                                  "DATE",
                                  "DESTINATION",
                                  "PARTNER",
                                  "MODE",
                                  "WT.",
                                  "PCS",
                                  "AMOUNT",
                                ].map(
                                  (heading) => (
                                    <th
                                      key={heading}
                                      className="border-b border-gray-200 px-3 py-3 text-[10px] font-bold tracking-wide"
                                    >
                                      {heading}
                                    </th>
                                  ),
                                )}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-100">
                              <tr>
                                <td className="px-2 py-2">
                                  <input
                                    required
                                    id="docket-no-input"
                                    aria-label="Docket number"
                                    onKeyDown={handleDocketInputKeyDown}
                                    value={form.docketNo}
                                    onChange={(event) =>
                                      updateDocketField("docketNo", event.target.value)
                                    }
                                    placeholder="Docket #"
                                    className="w-full min-w-0 rounded-md border border-gray-300 px-2 py-2 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                                  />
                                </td>
                                <td className="px-2 py-2">
                                  <input
                                    required
                                    aria-label="Consignment date"
                                    onKeyDown={handleDocketInputKeyDown}
                                    type="date"
                                    value={form.date}
                                    onChange={(event) => updateDocketField("date", event.target.value)}
                                    className="w-full min-w-0 rounded-md border border-gray-300 px-1.5 py-2 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                                  />
                                </td>
                                <td className="px-2 py-2">
                                  <Select<CityOption, false>
                                    required
                                    inputId="destination-input"
                                    aria-label="Destination"
                                    options={cityOptions}
                                    isSearchable
                                    value={
                                      cityOptions.find(
                                        (option) => option.value === form.destination,
                                      ) ?? null
                                    }
                                    onChange={(option) =>
                                      updateDocketField("destination", option?.value ?? "")
                                    }
                                    onKeyDown={(event) => {
                                      if (
                                        event.key === "Enter" &&
                                        event.target instanceof HTMLElement &&
                                        event.target.getAttribute("aria-expanded") === "true"
                                      ) {
                                        return;
                                      }
                                      handleDocketInputKeyDown(event);
                                    }}
                                    styles={destinationSelectStyles}
                                    placeholder="Destination"
                                  />
                                </td>
                                <td className="px-2 py-2">
                                  <select
                                    required
                                    aria-label="Courier partner"
                                    onKeyDown={handleDocketSelectKeyDown}
                                    value={form.courier}
                                    onChange={(event) =>
                                      updateDocketField("courier", event.target.value)
                                    }
                                    className="w-full min-w-0 rounded-md border border-gray-300 bg-white px-1.5 py-2 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                                  >
                                    <option value="" disabled>
                                      Partner
                                    </option>
                                    {courierPartners.map((partner) => (
                                      <option key={partner.id} value={partner.name}>
                                        {partner.name}
                                      </option>
                                    ))}
                                  </select>
                                </td>
                                <td className="px-2 py-2">
                                  <select
                                    required
                                    aria-label="Mode"
                                    onKeyDown={handleDocketSelectKeyDown}
                                    value={form.mode}
                                    onChange={(event) => updateDocketField("mode", event.target.value)}
                                    className="w-full min-w-0 rounded-md border border-gray-300 bg-white px-1.5 py-2 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                                  >
                                    <option value="" disabled>
                                      Mode
                                    </option>
                                    <option value="Surface">Surface</option>
                                    <option value="Air">Air</option>
                                  </select>
                                </td>
                                <td className="px-2 py-2">
                                  <input
                                    required
                                    aria-label="Weight"
                                    onKeyDown={handleDocketInputKeyDown}
                                    type="number"
                                    min="0"
                                    step="any"
                                    value={form.weight}
                                    onChange={(event) =>
                                      updateDocketField("weight", event.target.value)
                                    }
                                    placeholder="0.00"
                                    className="w-full min-w-0 rounded-md border border-gray-300 px-1.5 py-2 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                                  />
                                </td>
                                <td className="px-2 py-2">
                                  <input
                                    required
                                    aria-label="Pieces"
                                    onKeyDown={handleDocketInputKeyDown}
                                    type="number"
                                    min="1"
                                    step="1"
                                    value={form.pieces}
                                    onChange={(event) =>
                                      updateDocketField("pieces", event.target.value)
                                    }
                                    placeholder="1"
                                    className="w-full min-w-0 rounded-md border border-gray-300 px-1.5 py-2 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                                  />
                                </td>
                                <td className="px-2 py-2">
                                  <input
                                    required
                                    aria-label={
                                      invoiceType === "GST" ? "Amount including GST" : "Amount"
                                    }
                                    onKeyDown={handleDocketInputKeyDown}
                                    type="number"
                                    min="0"
                                    step="0.01"
                                    value={form.amount}
                                    onChange={(event) =>
                                      updateDocketField("amount", event.target.value)
                                    }
                                    placeholder="₹0.00"
                                    className="w-full min-w-0 rounded-md border border-gray-300 px-1.5 py-2 text-xs outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                                  />
                                </td>
                              </tr>
                              {dockets.map((docket, index) => (
                                <tr key={`${docket.docketNo}-${index}`} className="text-gray-600">
                                  <td className="whitespace-nowrap px-3 py-3 font-medium text-gray-900">
                                    {docket.docketNo}
                                  </td>
                                  <td className="whitespace-nowrap px-3 py-3">{docket.date}</td>
                                  <td className="px-3 py-3">{docket.destination}</td>
                                  <td className="px-3 py-3">{docket.courier}</td>
                                  <td className="px-3 py-3">{docket.mode}</td>
                                  <td className="px-3 py-3">{docket.weight}</td>
                                  <td className="px-3 py-3">{docket.pieces}</td>
                                  <td className="whitespace-nowrap px-3 py-3 text-right">
                                    {formatCurrency(docket.amount)}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                      </form>

                      <div className="border-t border-gray-200 px-5 py-5 sm:px-6">
                        <label className="block w-full text-sm font-medium text-gray-700">
                          Terms &amp; Conditions
                          <textarea
                            rows={4}
                            value={invoiceTerms}
                            onChange={(event) => setInvoiceTerms(event.target.value)}
                            className="mt-1.5 w-full resize-y rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                          />
                        </label>
                      </div>
                    </section>

                    <aside className="w-full rounded-2xl border border-slate-200 bg-white p-5 shadow-sm sm:p-6">
                      <div className="flex flex-wrap items-start justify-between gap-3">
                        <div>
                          <h2 className="text-lg font-bold text-slate-900">Billing Summary</h2>
                          {invoiceType === "GST" && (
                            <p className="mt-1 text-xs text-slate-500">
                              {isSameState ? "CGST + SGST" : "IGST"} at 18% · Based on{" "}
                              {selectedCustomerRecord?.state || "customer state"}
                            </p>
                          )}
                        </div>
                        <span
                          className={`rounded-full px-3 py-1 text-xs font-semibold ${
                            invoiceType === "GST"
                              ? "bg-blue-50 text-blue-700"
                              : "bg-slate-100 text-slate-700"
                          }`}
                        >
                          {invoiceType} Invoice
                        </span>
                      </div>

                      <div className="mt-5 grid gap-3 sm:grid-cols-3">
                        <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
                          <p className="text-xs font-medium text-slate-500">Subtotal</p>
                          <p className="mt-1 text-lg font-semibold tabular-nums text-slate-900">
                            {formatCurrency(subtotal)}
                          </p>
                        </div>
                        <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
                          <p className="text-xs font-medium text-slate-500">
                            {invoiceType === "GST" ? "GST Included" : "GST"}
                          </p>
                          <p className="mt-1 text-lg font-semibold tabular-nums text-slate-900">
                            {formatCurrency(totalGst)}
                          </p>
                          {invoiceType === "GST" && (
                            <p className="mt-1 text-xs text-slate-500">
                              {isSameState
                                ? `CGST ${formatCurrency(cgstTotal)} · SGST ${formatCurrency(sgstTotal)}`
                                : `IGST ${formatCurrency(igstTotal)}`}
                            </p>
                          )}
                        </div>
                        <div className="rounded-xl border border-blue-100 bg-blue-50 p-4">
                          <p className="text-xs font-medium text-blue-700">Grand Total</p>
                          <p className="mt-1 text-xl font-bold tabular-nums text-blue-950">
                            {formatCurrency(grandTotal)}
                          </p>
                        </div>
                      </div>
                      <p className="mt-3 text-right text-xs text-slate-500">
                        Round off: {formatCurrency(0)}
                      </p>

                      <div className="mt-5 flex flex-wrap items-end justify-between gap-4 border-t border-slate-100 pt-5">
                        <div className="grid w-full max-w-md flex-1 grid-cols-2 gap-3">
                          <label className="block text-xs font-semibold text-slate-600">
                            Status
                            <select
                              value={invoiceStatus}
                              onChange={(event) => setInvoiceStatus(event.target.value)}
                              className="mt-1.5 w-full rounded-md border border-slate-300 bg-white px-3 py-2.5 text-sm font-normal text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            >
                              <option value="Issued">Issued</option>
                              <option value="Draft">Draft</option>
                            </select>
                          </label>
                          <label className="block text-xs font-semibold text-slate-600">
                            Payment
                            <select
                              value={paymentStatus}
                              onChange={(event) => setPaymentStatus(event.target.value)}
                              className="mt-1.5 w-full rounded-md border border-slate-300 bg-white px-3 py-2.5 text-sm font-normal text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            >
                              <option value="Unpaid">Unpaid</option>
                              <option value="Paid">Paid</option>
                            </select>
                          </label>
                        </div>

                        {(invoiceSuccess || invoiceError) && (
                          <div className="w-full">
                            {invoiceSuccess && (
                              <p role="status" className="rounded-md bg-green-50 p-3 text-sm text-green-800">
                                {invoiceSuccess}
                              </p>
                            )}
                            {invoiceError && (
                              <p role="alert" className="rounded-md bg-red-50 p-3 text-sm text-red-700">
                                {invoiceError}
                              </p>
                            )}
                          </div>
                        )}
                        <button
                          type="button"
                          onClick={createInvoice}
                          className="ml-auto rounded-md bg-blue-700 px-6 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-200"
                        >
                          {editingInvoiceId ? "Update invoice" : "Create invoice"}
                        </button>
                      </div>
                    </aside>
                  </div>
                </section>
                ) : (
                  <section className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
                    <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 px-5 py-5 sm:px-6">
                      <div>
                        <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500">
                          BILLING
                        </p>
                        <h1 className="mt-1 text-2xl font-bold tracking-tight text-slate-950">
                          List of Invoices
                        </h1>
                        <p className="mt-1 text-sm text-slate-500">
                          {savedInvoices.length} saved invoice{savedInvoices.length === 1 ? "" : "s"}
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          startNewInvoice();
                        }}
                        className="inline-flex items-center gap-2 rounded-md bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-100"
                      >
                        <span className="text-lg leading-none">+</span>
                        New Invoice
                      </button>
                    </div>

                    <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 px-5 py-4 sm:px-6">
                      <nav className="flex flex-wrap gap-2" aria-label="Filter invoices">
                        {(
                          [
                            ["all", "All Invoices"],
                            ["GST", "GST Invoices"],
                            ["Non-GST", "Non-GST Invoices"],
                          ] as const
                        ).map(([filter, label]) => (
                          <button
                            key={filter}
                            type="button"
                            aria-pressed={invoiceFilter === filter}
                            onClick={() => setInvoiceFilter(filter)}
                            className={`rounded-full px-4 py-2 text-sm font-semibold transition ${
                              invoiceFilter === filter
                                ? "bg-blue-600 text-white shadow-sm"
                                : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                            }`}
                          >
                            {label}
                          </button>
                        ))}
                      </nav>
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={exportFilteredInvoices}
                          disabled={filteredInvoices.length === 0}
                          className="inline-flex items-center gap-2 rounded-md border border-slate-300 bg-white px-3.5 py-2 text-sm font-semibold text-slate-700 transition hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-50"
                        >
                          <svg
                            aria-hidden="true"
                            className="h-4 w-4"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="1.8"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <path d="M12 3v12m0 0 4-4m-4 4-4-4" />
                            <path d="M5 17v3h14v-3" />
                          </svg>
                          Export JSON
                        </button>
                        <button
                          type="button"
                          onClick={exportFilteredInvoicesToExcel}
                          disabled={filteredInvoices.length === 0}
                          className="inline-flex items-center gap-2 rounded-md bg-emerald-600 px-3.5 py-2 text-sm font-semibold text-white transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:opacity-50"
                        >
                          <svg
                            aria-hidden="true"
                            className="h-4 w-4"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="1.8"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <path d="M4 4h16v16H4z" />
                            <path d="m8 8 8 8m0-8-8 8M4 10h16M4 14h16" />
                          </svg>
                          Export Excel
                        </button>
                      </div>
                    </div>

                    {invoiceError && (
                      <p role="alert" className="mx-5 mt-4 rounded-md bg-red-50 p-3 text-sm text-red-700 sm:mx-6">
                        {invoiceError}
                      </p>
                    )}
                    {invoiceSuccess && (
                      <p role="status" className="mx-5 mt-4 rounded-md bg-green-50 p-3 text-sm text-green-800 sm:mx-6">
                        {invoiceSuccess}
                      </p>
                    )}

                    <div className="overflow-x-auto">
                      <table className="w-full min-w-[1100px] border-collapse text-left text-sm">
                        <thead className="bg-slate-50 text-slate-600">
                          <tr>
                            {[
                              "Invoice No.",
                              "Date",
                              "Customer Name",
                              "Invoice Type",
                              "GST Type",
                              "Total Amount",
                              "Status",
                              "Actions",
                            ].map((heading) => (
                              <th
                                key={heading}
                                className="border-b border-slate-200 px-5 py-3 text-xs font-semibold uppercase tracking-wide"
                              >
                                {heading}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {isLoadingData ? (
                            <tr>
                              <td colSpan={8} className="px-5 py-14 text-center text-slate-500">
                                <span role="status" className="inline-flex items-center gap-2">
                                  <span className="h-4 w-4 animate-spin rounded-full border-2 border-slate-300 border-t-blue-600" />
                                  Loading invoices...
                                </span>
                              </td>
                            </tr>
                          ) : filteredInvoices.length === 0 ? (
                            <tr>
                              <td colSpan={8} className="px-5 py-14 text-center">
                                <p className="font-medium text-slate-700">
                                  {savedInvoices.length === 0
                                    ? "No invoices yet"
                                    : "No invoices match this filter"}
                                </p>
                                <p className="mt-1 text-sm text-slate-500">
                                  {savedInvoices.length === 0
                                    ? "Create your first invoice to see it listed here."
                                    : "Choose another invoice type to view saved invoices."}
                                </p>
                              </td>
                            </tr>
                          ) : (
                            filteredInvoices.map((invoice) => (
                              <tr key={invoice._id} className="transition hover:bg-slate-50">
                                <td className="whitespace-nowrap px-5 py-4 font-semibold text-slate-900">
                                  {invoice.invoiceNo}
                                </td>
                                <td className="whitespace-nowrap px-5 py-4 text-slate-600">
                                  {invoice.date}
                                </td>
                                <td className="px-5 py-4 font-medium text-slate-800">
                                  {invoice.customer.name}
                                </td>
                                <td className="px-5 py-4">
                                  <span
                                    className={`inline-flex rounded-full px-2.5 py-1 text-xs font-medium ${
                                      invoice.invoiceType === "Non-GST"
                                        ? "bg-slate-100 text-slate-700"
                                        : "bg-blue-50 text-blue-700"
                                    }`}
                                  >
                                    {invoice.invoiceType ?? "GST"}
                                  </span>
                                </td>
                                <td className="px-5 py-4 text-slate-600">
                                  {invoice.invoiceType === "Non-GST" ? "—" : invoice.gstType}
                                </td>
                                <td className="whitespace-nowrap px-5 py-4 font-semibold text-slate-900">
                                  {formatCurrency(invoice.totalAmount)}
                                </td>
                                <td className="px-5 py-4">
                                  <span
                                    className={`inline-flex rounded-full px-2.5 py-1 text-xs font-medium ${
                                      invoice.status === "Draft"
                                        ? "bg-amber-50 text-amber-700"
                                        : "bg-emerald-50 text-emerald-700"
                                    }`}
                                  >
                                    {invoice.status}
                                  </span>
                                </td>
                                <td className="px-5 py-4">
                                  <div className="flex items-center gap-2">
                                    <button
                                      type="button"
                                      onClick={() => editInvoice(invoice)}
                                      className="rounded-md border border-blue-200 bg-white px-3 py-1.5 text-xs font-semibold text-blue-700 transition hover:bg-blue-50"
                                    >
                                      Edit
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => printSavedInvoice(invoice)}
                                      className="rounded-md border border-emerald-200 bg-white px-3 py-1.5 text-xs font-semibold text-emerald-700 transition hover:bg-emerald-50"
                                    >
                                      Download
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => printSavedInvoice(invoice)}
                                      className="rounded-md border border-orange-200 bg-white px-3 py-1.5 text-xs font-semibold text-orange-700 transition hover:bg-orange-50"
                                    >
                                      Print
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => deleteInvoice(invoice._id)}
                                      className="rounded-md border border-red-200 bg-white px-3 py-1.5 text-xs font-semibold text-red-600 transition hover:bg-red-50"
                                    >
                                      Delete
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </section>
                )
              )}

              {activeTab === "customers" && (
                <section className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
                    <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 px-5 py-5 sm:px-6">
                      <div>
                        <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500">
                          MASTERS
                        </p>
                        <h1 className="mt-1 text-2xl font-bold tracking-tight text-slate-950">
                          Customers
                        </h1>
                        <p className="mt-1 text-sm text-slate-500">
                          {customers.length} customer{customers.length === 1 ? "" : "s"}
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          setCustomerForm({ ...emptyCustomerForm });
                          setCustomerError("");
                          setIsAddingCustomer(true);
                        }}
                        className="inline-flex items-center gap-2 rounded-md bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-100"
                      >
                        <span className="text-lg leading-none">+</span>
                        Add Customer
                      </button>
                    </div>

                    {customerError && !isAddingCustomer && (
                      <p role="alert" className="mx-5 mt-4 rounded-md bg-red-50 p-3 text-sm text-red-700 sm:mx-6">
                        {customerError}
                      </p>
                    )}

                    <div className="overflow-x-auto">
                      <table className="w-full min-w-[850px] border-collapse text-left text-sm">
                        <thead className="bg-slate-50 text-slate-600">
                          <tr>
                            {["Name", "Phone", "State", "GSTIN", "Actions"].map((heading) => (
                              <th
                                key={heading}
                                className="border-b border-slate-200 px-5 py-3 text-xs font-semibold uppercase tracking-wide"
                              >
                                {heading}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {isLoadingData ? (
                            <tr>
                              <td colSpan={5} className="px-5 py-14 text-center text-slate-500">
                                <span role="status" className="inline-flex items-center gap-2">
                                  <span className="h-4 w-4 animate-spin rounded-full border-2 border-slate-300 border-t-blue-600" />
                                  Loading customers...
                                </span>
                              </td>
                            </tr>
                          ) : customers.length === 0 ? (
                            <tr>
                              <td colSpan={5} className="px-5 py-14 text-center">
                                <p className="font-medium text-slate-700">No customers yet</p>
                                <p className="mt-1 text-sm text-slate-500">
                                  Add a customer to use them on invoices.
                                </p>
                              </td>
                            </tr>
                          ) : (
                            customers.map((customer, index) => (
                              <tr
                                key={customer._id ?? `${customer.phone}-${index}`}
                                className="transition hover:bg-slate-50"
                              >
                                <td className="px-5 py-4 font-semibold text-slate-900">
                                  {customer.name}
                                </td>
                                <td className="whitespace-nowrap px-5 py-4 text-slate-600">
                                  {customer.phone}
                                </td>
                                <td className="px-5 py-4 text-slate-600">{customer.state}</td>
                                <td className="px-5 py-4 text-slate-600">
                                  {customer.gstin || "—"}
                                </td>
                                <td className="px-5 py-4">
                                  <button
                                    type="button"
                                    onClick={() => deleteCustomer(customer)}
                                    className="rounded-md border border-red-200 bg-white px-3 py-1.5 text-xs font-semibold text-red-600 transition hover:bg-red-50"
                                  >
                                    Delete
                                  </button>
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                </section>
              )}

              {activeTab === "couriers" && (
                <section className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
                  <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 px-5 py-5 sm:px-6">
                    <div>
                      <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500">
                        MASTERS
                      </p>
                      <h1 className="mt-1 text-3xl font-bold tracking-tight text-slate-950">
                        Courier Partners
                      </h1>
                      <p className="mt-1 text-sm text-slate-500">
                        {courierPartners.length} saved partner{courierPartners.length === 1 ? "" : "s"}
                      </p>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        setCourierError("");
                        setIsAddingCourierPartner((isAdding) => !isAdding);
                      }}
                      className="inline-flex items-center gap-2 rounded-md bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700"
                    >
                      <span className="text-lg leading-none">+</span>
                      Add Partner
                    </button>
                  </div>

                  {isAddingCourierPartner && (
                    <form
                      onSubmit={addCourierPartner}
                      className="border-b border-slate-200 bg-slate-50 p-5 sm:p-6"
                    >
                      <div className="grid gap-4 sm:grid-cols-2">
                        {(
                          [
                            ["name", "Name", true],
                            ["contactPerson", "Contact Person", false],
                            ["phone", "Phone", false],
                          ] as const
                        ).map(([field, label, required]) => (
                          <label
                            key={field}
                            className="block text-xs font-semibold uppercase tracking-wide text-slate-500"
                          >
                            {label}
                            <input
                              required={required}
                              type={field === "phone" ? "tel" : "text"}
                              value={courierPartnerForm[field]}
                              onChange={(event) =>
                                setCourierPartnerForm((current) => ({
                                  ...current,
                                  [field]: event.target.value,
                                }))
                              }
                              className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            />
                          </label>
                        ))}
                        <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500 sm:col-span-2">
                          Notes
                          <textarea
                            rows={3}
                            value={courierPartnerForm.notes}
                            onChange={(event) =>
                              setCourierPartnerForm((current) => ({
                                ...current,
                                notes: event.target.value,
                              }))
                            }
                            className="mt-1.5 w-full resize-y rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                          />
                        </label>
                      </div>
                      {courierError && (
                        <p role="alert" className="mt-4 rounded-md bg-red-50 p-3 text-sm text-red-700">
                          {courierError}
                        </p>
                      )}
                      <div className="mt-4 flex justify-end gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            setIsAddingCourierPartner(false);
                            setCourierPartnerForm({ ...emptyCourierPartnerForm });
                            setCourierError("");
                          }}
                          className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
                        >
                          Cancel
                        </button>
                        <button
                          type="submit"
                          className="rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-700"
                        >
                          Save Partner
                        </button>
                      </div>
                    </form>
                  )}

                  {courierError && !isAddingCourierPartner && (
                    <p role="alert" className="mx-5 mt-4 rounded-md bg-red-50 p-3 text-sm text-red-700 sm:mx-6">
                      {courierError}
                    </p>
                  )}

                  <div className="overflow-x-auto">
                    <table className="w-full min-w-[800px] border-collapse text-left text-sm">
                      <thead className="bg-slate-50 text-slate-600">
                        <tr>
                          {["Name", "Contact Person", "Phone", "Notes", "Actions"].map((heading) => (
                            <th
                              key={heading}
                              className="border-b border-slate-200 px-5 py-3 text-xs font-semibold uppercase tracking-wide"
                            >
                              {heading}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {courierPartners.length === 0 ? (
                          <tr>
                            <td colSpan={5} className="px-5 py-14 text-center">
                              <p className="font-medium text-slate-700">No courier partners yet</p>
                              <p className="mt-1 text-sm text-slate-500">
                                Add a partner to select it when creating invoices.
                              </p>
                            </td>
                          </tr>
                        ) : (
                          courierPartners.map((partner) => (
                            <tr key={partner.id} className="transition hover:bg-slate-50">
                              <td className="px-5 py-4 font-semibold text-slate-900">{partner.name}</td>
                              <td className="px-5 py-4 text-slate-600">{partner.contactPerson || "—"}</td>
                              <td className="whitespace-nowrap px-5 py-4 text-slate-600">{partner.phone || "—"}</td>
                              <td className="max-w-sm whitespace-pre-wrap px-5 py-4 text-slate-600">{partner.notes || "—"}</td>
                              <td className="px-5 py-4">
                                <button
                                  type="button"
                                  onClick={() => deleteCourierPartner(partner.id)}
                                  className="rounded-md border border-red-200 bg-white px-3 py-1.5 text-xs font-semibold text-red-600 transition hover:bg-red-50"
                                >
                                  Delete
                                </button>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}

              {activeTab === "reports" && (
                <section className="space-y-6">
                  <div>
                    <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500">
                      ANALYTICS
                    </p>
                    <h1 className="mt-1 text-3xl font-bold tracking-tight text-slate-950">
                      Reports
                    </h1>
                    <p className="mt-1 text-sm text-slate-500">
                      Review invoice sales and GST for a selected date range.
                    </p>
                  </div>

                  <div className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
                    <div className="grid items-end gap-4 sm:grid-cols-[1fr_1fr_auto]">
                      <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500">
                        Start Date
                        <input
                          type="date"
                          value={reportStartDate}
                          onChange={(event) => {
                            setReportStartDate(event.target.value);
                            setReportError("");
                          }}
                          className="mt-1.5 w-full rounded-md border border-slate-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                        />
                      </label>
                      <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500">
                        End Date
                        <input
                          type="date"
                          value={reportEndDate}
                          onChange={(event) => {
                            setReportEndDate(event.target.value);
                            setReportError("");
                          }}
                          className="mt-1.5 w-full rounded-md border border-slate-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-slate-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                        />
                      </label>
                      <button
                        type="button"
                        onClick={generateReport}
                        className="rounded-md bg-blue-600 px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-100"
                      >
                        Generate Report
                      </button>
                    </div>
                    {reportError && (
                      <p role="alert" className="mt-4 rounded-md bg-red-50 p-3 text-sm text-red-700">
                        {reportError}
                      </p>
                    )}
                  </div>

                  <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
                    {[
                      ["Total Invoices", String(reportInvoices.length)],
                      ["Total Sales", formatCurrency(roundCurrencyAmount(reportTotals.totalSales))],
                      ["Total GST Collected", formatCurrency(roundCurrencyAmount(reportTotals.gstCollected))],
                      ["Total Non-GST Sales", formatCurrency(roundCurrencyAmount(reportTotals.nonGstSales))],
                    ].map(([label, value]) => (
                      <article
                        key={label}
                        className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm"
                      >
                        <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                          {label}
                        </p>
                        <p className="mt-3 text-xl font-bold text-slate-950">{value}</p>
                      </article>
                    ))}
                  </div>

                  <section className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
                    <div className="border-b border-slate-100 px-5 py-4 sm:px-6">
                      <h2 className="font-semibold text-slate-900">Invoice Breakdown</h2>
                      <p className="mt-1 text-sm text-slate-500">
                        {generatedReportRange.startDate} to {generatedReportRange.endDate}
                      </p>
                    </div>
                    <div className="max-h-[600px] overflow-auto">
                      <table className="w-full min-w-[1000px] border-collapse text-left text-sm">
                        <thead className="sticky top-0 bg-slate-50 text-slate-600">
                          <tr>
                            {[
                              "Invoice No",
                              "Date",
                              "Customer Name",
                              "Type",
                              "Taxable Amount",
                              "GST Amount",
                              "Total Amount",
                            ].map((heading) => (
                              <th
                                key={heading}
                                className="border-b border-slate-200 px-5 py-3 text-xs font-semibold uppercase tracking-wide"
                              >
                                {heading}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {reportInvoices.length === 0 ? (
                            <tr>
                              <td colSpan={7} className="px-5 py-12 text-center text-slate-500">
                                No invoices found in this date range.
                              </td>
                            </tr>
                          ) : (
                            reportInvoices.map((invoice) => {
                              const invoiceType = invoice.invoiceType ?? "GST";
                              const gstAmount =
                                invoiceType === "Non-GST"
                                  ? 0
                                  : invoice.cgstTotal + invoice.sgstTotal + invoice.igstTotal;

                              return (
                                <tr key={invoice._id} className="transition hover:bg-slate-50">
                                  <td className="whitespace-nowrap px-5 py-4 font-semibold text-slate-900">
                                    {invoice.invoiceNo}
                                  </td>
                                  <td className="whitespace-nowrap px-5 py-4 text-slate-600">
                                    {invoice.date}
                                  </td>
                                  <td className="px-5 py-4 font-medium text-slate-800">
                                    {invoice.customer.name}
                                  </td>
                                  <td className="px-5 py-4 text-slate-600">{invoiceType}</td>
                                  <td className="whitespace-nowrap px-5 py-4 text-right tabular-nums text-slate-700">
                                    {formatCurrency(invoice.subtotal)}
                                  </td>
                                  <td className="whitespace-nowrap px-5 py-4 text-right tabular-nums text-slate-700">
                                    {formatCurrency(gstAmount)}
                                  </td>
                                  <td className="whitespace-nowrap px-5 py-4 text-right font-semibold tabular-nums text-slate-900">
                                    {formatCurrency(invoice.totalAmount)}
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </section>
                </section>
              )}

              {activeTab === "settings" && (
                <section className="max-w-5xl">
                  <div className="mb-6">
                    <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500">
                      PREFERENCES
                    </p>
                    <h1 className="mt-1 text-3xl font-bold tracking-tight text-slate-950">
                      Settings
                    </h1>
                    <p className="mt-1 text-sm text-slate-500">
                      Manage the company and bank information shown on invoices.
                    </p>
                  </div>
                  <form onSubmit={saveCompanySettings} className="space-y-6">
                    <section className="rounded-lg border border-slate-200 bg-white shadow-sm">
                      <div className="border-b border-slate-100 px-5 py-4 sm:px-6">
                        <h2 className="font-semibold text-slate-900">Company Profile</h2>
                        <p className="mt-1 text-sm text-slate-500">
                          These details appear on printed invoices and determine the company state for GST.
                        </p>
                      </div>
                      <div className="grid gap-4 p-5 sm:grid-cols-2 sm:p-6">
                        {([
                          ["name", "Company Name", "text"],
                          ["gstin", "GSTIN", "text"],
                          ["phone", "Phone", "tel"],
                          ["email", "Email", "email"],
                          ["address", "Address", "text"],
                          ["city", "City", "text"],
                          ["state", "State", "text"],
                          ["pincode", "Pincode", "text"],
                        ] as const).map(([field, label, type]) => (
                          <label
                            key={field}
                            className={`block text-xs font-semibold uppercase tracking-wide text-slate-500 ${
                              field === "address" ? "sm:col-span-2" : ""
                            }`}
                          >
                            {label}
                            <input
                              type={type}
                              value={companySettings[field]}
                              onChange={(event) => updateCompanyField(field, event.target.value)}
                              className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            />
                          </label>
                        ))}

                        <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500 sm:col-span-2">
                          UPI ID
                          <input
                            type="text"
                            value={companySettings.upiId}
                            onChange={(event) => updateCompanyField("upiId", event.target.value)}
                            placeholder="yourname@upi"
                            autoComplete="off"
                            className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                          />
                        </label>
                        {companySettings.upiId.trim() && (
                          <div className="flex items-center gap-4 rounded-md border border-slate-200 bg-slate-50 p-4 sm:col-span-2">
                            <QRCode
                              value={`upi://pay?pa=${encodeURIComponent(
                                companySettings.upiId.trim(),
                              )}&pn=${encodeURIComponent(companySettings.name)}&cu=INR`}
                              size={112}
                              bgColor="#ffffff"
                              fgColor="#000000"
                              aria-label="UPI payment QR code preview"
                            />
                            <div>
                              <p className="text-sm font-semibold text-slate-800">
                                UPI payment QR preview
                              </p>
                              <p className="mt-1 text-xs text-slate-500">
                                Scan to pay {companySettings.name || "your company"} via UPI.
                              </p>
                            </div>
                          </div>
                        )}

                        <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500 sm:col-span-2">
                          Logo Upload
                          <input
                            type="file"
                            accept="image/*"
                            onChange={(event) => readCompanyImage("logo", event.target.files?.[0])}
                            className="mt-1.5 block w-full rounded-md border border-gray-300 bg-white text-sm font-normal normal-case text-gray-700 file:mr-4 file:border-0 file:bg-slate-50 file:px-4 file:py-2.5 file:text-sm file:font-semibold file:text-slate-700 hover:file:bg-slate-100"
                          />
                          <span className="mt-1 block font-normal normal-case tracking-normal text-slate-400">
                            Image files are embedded in invoice printouts.
                          </span>
                        </label>
                        {companySettings.logo && (
                          <div className="flex items-center gap-4 sm:col-span-2">
                            {/* eslint-disable-next-line @next/next/no-img-element */}
                            <img
                              src={companySettings.logo}
                              alt="Company logo preview"
                              className="h-16 max-w-40 rounded border border-slate-200 bg-white object-contain p-2"
                            />
                            <button
                              type="button"
                              onClick={() => updateCompanyField("logo", "")}
                              className="text-sm font-medium text-red-600 hover:text-red-700"
                            >
                              Remove logo
                            </button>
                          </div>
                        )}

                        <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500 sm:col-span-2">
                          Authorized Signature Upload
                          <input
                            type="file"
                            accept="image/*"
                            onChange={(event) =>
                              readCompanyImage("signature", event.target.files?.[0])
                            }
                            className="mt-1.5 block w-full rounded-md border border-gray-300 bg-white text-sm font-normal normal-case text-gray-700 file:mr-4 file:border-0 file:bg-slate-50 file:px-4 file:py-2.5 file:text-sm file:font-semibold file:text-slate-700 hover:file:bg-slate-100"
                          />
                          <span className="mt-1 block font-normal normal-case tracking-normal text-slate-400">
                            The uploaded signature appears on printed invoices.
                          </span>
                        </label>
                        {companySettings.signature && (
                          <div className="flex items-center gap-4 sm:col-span-2">
                            {/* eslint-disable-next-line @next/next/no-img-element */}
                            <img
                              src={companySettings.signature}
                              alt="Authorized signature preview"
                              className="h-16 max-w-48 rounded border border-slate-200 bg-white object-contain p-2"
                            />
                            <button
                              type="button"
                              onClick={() => updateCompanyField("signature", "")}
                              className="text-sm font-medium text-red-600 hover:text-red-700"
                            >
                              Remove signature
                            </button>
                          </div>
                        )}
                      </div>
                    </section>

                    <section className="rounded-lg border border-slate-200 bg-white shadow-sm">
                      <div className="border-b border-slate-100 px-5 py-4 sm:px-6">
                        <h2 className="font-semibold text-slate-900">Bank Details</h2>
                        <p className="mt-1 text-sm text-slate-500">
                          Payment information included with the invoice.
                        </p>
                      </div>
                      <div className="grid gap-4 p-5 sm:grid-cols-2 sm:p-6">
                        {([
                          ["bankName", "Bank Name"],
                          ["accountNumber", "Account Number"],
                          ["ifscCode", "IFSC Code"],
                          ["branchName", "Branch Name"],
                        ] as const).map(([field, label]) => (
                          <label
                            key={field}
                            className="block text-xs font-semibold uppercase tracking-wide text-slate-500"
                          >
                            {label}
                            <input
                              type="text"
                              value={companySettings[field]}
                              onChange={(event) => updateCompanyField(field, event.target.value)}
                              className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                            />
                          </label>
                        ))}
                      </div>
                    </section>

                    {settingsError && (
                      <p role="alert" className="rounded-md bg-red-50 p-3 text-sm text-red-700">
                        {settingsError}
                      </p>
                    )}
                    {settingsSuccess && (
                      <p role="status" className="rounded-md bg-green-50 p-3 text-sm text-green-800">
                        {settingsSuccess}
                      </p>
                    )}
                    <div className="flex justify-end">
                      <button
                        type="submit"
                        className="rounded-md bg-blue-700 px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-200"
                      >
                        Save Settings
                      </button>
                    </div>
                  </form>

                  <section className="mt-6 rounded-lg border border-slate-200 bg-white shadow-sm">
                    <div className="border-b border-slate-100 px-5 py-4 sm:px-6">
                      <h2 className="font-semibold text-slate-900">Account Security</h2>
                      <p className="mt-1 text-sm text-slate-500">
                        Update the email and password used to sign in.
                      </p>
                    </div>
                    <form onSubmit={saveAccountSecurity} className="space-y-4 p-5 sm:p-6">
                      <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500">
                        Login Email
                        <input
                          required
                          type="email"
                          autoComplete="username"
                          value={loginEmail}
                          onChange={(event) => {
                            setLoginEmail(event.target.value);
                            setSecurityError("");
                            setSecuritySuccess("");
                          }}
                          className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                        />
                      </label>
                      <div className="grid gap-4 sm:grid-cols-2">
                        <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500">
                          New Password
                          <input
                            required
                            type="password"
                            autoComplete="new-password"
                            minLength={8}
                            value={newPassword}
                            onChange={(event) => {
                              setNewPassword(event.target.value);
                              setSecurityError("");
                              setSecuritySuccess("");
                            }}
                            className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                          />
                        </label>
                        <label className="block text-xs font-semibold uppercase tracking-wide text-slate-500">
                          Confirm Password
                          <input
                            required
                            type="password"
                            autoComplete="new-password"
                            minLength={8}
                            value={confirmPassword}
                            onChange={(event) => {
                              setConfirmPassword(event.target.value);
                              setSecurityError("");
                              setSecuritySuccess("");
                            }}
                            className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                          />
                        </label>
                      </div>
                      {securityError && (
                        <p role="alert" className="rounded-md bg-red-50 p-3 text-sm text-red-700">
                          {securityError}
                        </p>
                      )}
                      {securitySuccess && (
                        <p role="status" className="rounded-md bg-green-50 p-3 text-sm text-green-800">
                          {securitySuccess}
                        </p>
                      )}
                      <div className="flex justify-end">
                        <button
                          type="submit"
                          className="rounded-md bg-blue-700 px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-200"
                        >
                          Update Credentials
                        </button>
                      </div>
                    </form>
                  </section>
                </section>
              )}
            </div>
          </div>
        </div>
      )}

      {isLoggedIn && isAddingCustomer && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center overflow-y-auto bg-black/50 p-4 print:hidden sm:p-6"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) {
              setIsAddingCustomer(false);
              setCustomerError("");
            }
          }}
        >
          <section
            role="dialog"
            aria-modal="true"
            aria-labelledby="add-customer-title"
            className="my-auto w-full max-w-2xl overflow-hidden rounded-xl bg-white shadow-2xl"
          >
            <header className="flex items-start justify-between border-b border-gray-200 px-6 py-5">
              <div>
                <h2 id="add-customer-title" className="text-xl font-bold text-gray-900">
                  Add customer
                </h2>
                <p className="mt-1.5 text-sm text-gray-500">
                  Master record used for billing. GSTIN and State Code drive GST type.
                </p>
              </div>
              <button
                type="button"
                aria-label="Close add customer modal"
                onClick={() => {
                  setIsAddingCustomer(false);
                  setCustomerError("");
                }}
                className="ml-4 rounded-md p-1.5 text-gray-400 transition hover:bg-gray-100 hover:text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <svg
                  aria-hidden="true"
                  className="h-5 w-5"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                >
                  <path d="m18 6-12 12M6 6l12 12" />
                </svg>
              </button>
            </header>

            <form onSubmit={addCustomer}>
              <div className="max-h-[calc(100vh-230px)] overflow-y-auto px-6 py-5">
                <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    Name <span className="text-red-500">*</span>
                    <input
                      required
                      autoFocus
                      type="text"
                      value={customerForm.name}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, name: event.target.value }))
                      }
                      placeholder="Customer name"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>
                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    Contact Person
                    <input
                      type="text"
                      value={customerForm.contactPerson ?? ""}
                      onChange={(event) =>
                        setCustomerForm((current) => ({
                          ...current,
                          contactPerson: event.target.value,
                        }))
                      }
                      placeholder="Contact person"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>

                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    Phone
                    <input
                      type="tel"
                      value={customerForm.phone}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, phone: event.target.value }))
                      }
                      placeholder="Phone number"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>
                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    Email
                    <input
                      type="email"
                      value={customerForm.email ?? ""}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, email: event.target.value }))
                      }
                      placeholder="name@company.com"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>

                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    GSTIN
                    <input
                      type="text"
                      value={customerForm.gstin ?? ""}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, gstin: event.target.value }))
                      }
                      placeholder="GST identification number"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal uppercase tracking-normal text-gray-900 outline-none transition placeholder:normal-case placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>
                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    PAN
                    <input
                      type="text"
                      value={customerForm.pan ?? ""}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, pan: event.target.value }))
                      }
                      placeholder="PAN"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal uppercase tracking-normal text-gray-900 outline-none transition placeholder:normal-case placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>

                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500 sm:col-span-2">
                    Address
                    <textarea
                      rows={3}
                      value={customerForm.address}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, address: event.target.value }))
                      }
                      placeholder="Street address, building, area"
                      className="mt-1.5 w-full resize-y rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>

                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    City
                    <input
                      type="text"
                      value={customerForm.city ?? ""}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, city: event.target.value }))
                      }
                      placeholder="City"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>
                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    State <span className="text-red-500">*</span>
                    <select
                      required
                      value={customerForm.state}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, state: event.target.value }))
                      }
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    >
                      {indianStates.map((state) => (
                        <option key={state} value={state}>
                          {state}
                        </option>
                      ))}
                    </select>
                  </label>

                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    State Code
                    <input
                      type="text"
                      value={customerForm.stateCode ?? ""}
                      onChange={(event) =>
                        setCustomerForm((current) => ({
                          ...current,
                          stateCode: event.target.value,
                        }))
                      }
                      placeholder="State code"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>
                  <label className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                    Pincode
                    <input
                      type="text"
                      inputMode="numeric"
                      value={customerForm.pincode ?? ""}
                      onChange={(event) =>
                        setCustomerForm((current) => ({ ...current, pincode: event.target.value }))
                      }
                      placeholder="PIN code"
                      className="mt-1.5 w-full rounded-md border border-gray-300 bg-white px-3 py-2.5 text-sm font-normal normal-case tracking-normal text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                    />
                  </label>
                </div>

                {customerError && (
                  <p role="alert" className="mt-4 rounded-md bg-red-50 p-3 text-sm text-red-700">
                    {customerError}
                  </p>
                )}
              </div>

              <footer className="flex justify-end gap-3 border-t border-gray-200 bg-gray-50 px-6 py-4">
                <button
                  type="button"
                  onClick={() => {
                    setCustomerForm({ ...emptyCustomerForm });
                    setCustomerError("");
                    setIsAddingCustomer(false);
                  }}
                  className="rounded-md border border-gray-300 bg-white px-4 py-2.5 text-sm font-semibold text-gray-700 shadow-sm transition hover:bg-gray-50"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="rounded-md bg-blue-700 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-200"
                >
                  Save customer
                </button>
              </footer>
            </form>
          </section>
        </div>
      )}

      {selectedPrintInvoice && (
        <div
          ref={printInvoiceRef}
          id="print-invoice-container"
          className="mx-auto hidden w-full max-w-4xl bg-white p-8 font-sans text-slate-800 print:block print:w-[210mm] print:max-w-none print:p-4"
        >
          <header className="flex items-center justify-between gap-6 border-b border-gray-200 pb-5 print:gap-3 print:pb-2">
            <div className="flex items-center gap-4">
              {printCompanySettings.logo ? (
                <div className="mr-6 flex h-24 w-auto max-w-[180px] shrink-0 items-center justify-center overflow-hidden">
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img
                    src={printCompanySettings.logo}
                    alt={`${printCompanySettings.name || "Company"} logo`}
                    className="h-full w-full object-contain"
                    onError={(event) => {
                      event.currentTarget.style.display = "none";
                      event.currentTarget.nextElementSibling?.classList.remove("hidden");
                    }}
                  />
                  <span className="hidden text-center text-xs font-bold text-blue-700">
                    {printCompanySettings.name || "Company"}
                  </span>
                </div>
              ) : (
                <div className="flex h-20 w-28 shrink-0 items-center justify-center rounded-lg bg-blue-50 text-lg font-bold text-blue-700">
                  {(printCompanySettings.name || "AGC")
                    .split(/\s+/)
                    .slice(0, 2)
                    .map((word) => word[0])
                    .join("")}
                </div>
              )}
              <div>
                <h1 className="text-2xl font-bold tracking-tight text-slate-900">
                  {printCompanySettings.name || "AGC BILLING"}
                </h1>
                {printCompanySettings.gstin && (
                  <p className="mt-1 text-sm text-slate-600">GSTIN: {printCompanySettings.gstin}</p>
                )}
                {printCompanySettings.address && (
                  <p className="mt-1 text-xs text-slate-600">
                    {printCompanySettings.address}
                    {printCompanySettings.city ? `, ${printCompanySettings.city}` : ""}
                    {printCompanySettings.state ? `, ${printCompanySettings.state}` : ""}
                    {printCompanySettings.pincode ? ` - ${printCompanySettings.pincode}` : ""}
                  </p>
                )}
                {(printCompanySettings.phone || printCompanySettings.email) && (
                  <p className="mt-1 text-xs text-slate-600">
                    {[printCompanySettings.phone, printCompanySettings.email]
                      .filter(Boolean)
                      .join(" · ")}
                  </p>
                )}
              </div>
            </div>
            <div className="shrink-0 text-right">
              <p className="text-3xl font-bold tracking-tight text-blue-700">
                {printInvoiceType === "GST" ? "TAX INVOICE" : "INVOICE"}
              </p>
              <p className="mt-1 text-xs font-semibold uppercase tracking-[0.16em] text-slate-500">
                {selectedPrintInvoice.invoiceNo}
              </p>
            </div>
          </header>
          <section className="mt-6 grid grid-cols-2 gap-6 text-sm print:mt-3 print:gap-3">
            <div className="space-y-3">
              <div>
                <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">Invoice Date</p>
                <p className="mt-1 font-medium text-slate-900">{selectedPrintInvoice.date}</p>
              </div>
              {selectedPrintInvoice.dueDate && (
                <div>
                  <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">Due Date</p>
                  <p className="mt-1 font-medium text-slate-900">{selectedPrintInvoice.dueDate}</p>
                </div>
              )}
              <div>
                <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">SAC Code</p>
                <p className="mt-1 font-medium text-slate-900">996812</p>
              </div>
            </div>
            <div className="rounded-lg border border-gray-200 bg-gray-50 p-4">
              <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">Bill To</p>
              <div className="mt-2 space-y-1 text-slate-700">
                <p className="font-semibold text-slate-900">{selectedPrintInvoice.customer.name}</p>
                {selectedPrintInvoice.customer.address && <p>{selectedPrintInvoice.customer.address}</p>}
                <p>
                  {[
                    selectedPrintInvoice.customer.city,
                    selectedPrintInvoice.customer.state,
                    selectedPrintInvoice.customer.phone,
                  ]
                    .filter(Boolean)
                    .join(" · ")}
                </p>
                {selectedPrintInvoice.customer.gstin && (
                  <p>GSTIN: {selectedPrintInvoice.customer.gstin}</p>
                )}
              </div>
            </div>
          </section>
          <table className="mt-7 w-full table-fixed border-collapse text-left text-[11px] print:mt-3">
            <thead className="print:table-header-group">
              <tr className="bg-gray-100 text-[10px] uppercase tracking-wide text-slate-600">
                {[
                  "S.No.",
                  "Docket No.",
                  "Date",
                  "Destination",
                  "Courier",
                  "Wt.",
                  "Pcs",
                  "Taxable Value",
                ].map((heading) => (
                  <th key={heading} className="border border-gray-200 p-3 font-bold print:p-1">
                    {heading}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {printConsignments.length > 0 ? (
                printConsignments.map((docket, index) => {
                  const taxableValue =
                    selectedPrintInvoice.gstInclusive === true
                      ? docket.taxableValue ??
                        calculateInclusiveDocketTax(
                          docket.amount,
                          printGstType,
                          printInvoiceType,
                        ).taxableValue
                      : docket.amount;

                  return (
                    <tr key={`${docket.docketNo}-${index}`} className="print:break-inside-avoid">
                      <td className="border border-gray-200 p-3 text-center print:p-1">{index + 1}</td>
                      <td className="break-words border border-gray-200 p-3 print:p-1">{docket.docketNo}</td>
                      <td className="whitespace-nowrap border border-gray-200 p-3 print:p-1">{docket.date}</td>
                      <td className="break-words border border-gray-200 p-3 print:p-1">{docket.destination}</td>
                      <td className="break-words border border-gray-200 p-3 print:p-1">{docket.courier}</td>
                      <td className="border border-gray-200 p-3 print:p-1">{docket.weight}</td>
                      <td className="border border-gray-200 p-3 print:p-1">{docket.pieces}</td>
                      <td className="border border-gray-200 p-3 text-right tabular-nums print:p-1">
                        {formatCurrency(taxableValue)}
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td
                    colSpan={8}
                    className="border border-gray-200 p-4 text-center text-slate-500"
                  >
                    No dockets added.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
          <footer className="mt-8 h-auto print:mt-4">
            <div className="ml-auto w-64 max-w-full space-y-2 rounded-lg border border-gray-300 bg-gray-50 p-4 text-sm print:break-inside-avoid print:p-2">
              <div className="flex justify-between">
                <span>Subtotal</span>
                <span className="tabular-nums">{formatCurrency(printSubtotal)}</span>
              </div>
              {printInvoiceType === "GST" &&
                (selectedPrintInvoice.gstType === "CGST+SGST" ? (
                  <>
                    <div className="flex justify-between">
                      <span>CGST (9%)</span>
                      <span className="tabular-nums">{formatCurrency(printCgstTotal)}</span>
                    </div>
                    <div className="flex justify-between">
                      <span>SGST (9%)</span>
                      <span className="tabular-nums">{formatCurrency(printSgstTotal)}</span>
                    </div>
                  </>
                ) : (
                  <div className="flex justify-between">
                    <span>IGST (18%)</span>
                    <span className="tabular-nums">{formatCurrency(printIgstTotal)}</span>
                  </div>
                ))}
              <div className="flex justify-between border-t border-gray-300 pt-3 text-lg font-bold text-slate-900">
                <span>Grand Total</span>
                <span className="tabular-nums">{formatCurrency(printTotalAmount)}</span>
              </div>
            </div>
            {printCompanySettings.signature && (
              <div className="ml-auto mt-6 w-64 max-w-full text-center print:mt-3 print:break-inside-avoid">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img
                  src={printCompanySettings.signature}
                  alt={`${printCompanySettings.name || "Company"} authorized signature`}
                  className="mx-auto h-16 max-w-48 object-contain"
                />
                <div className="mt-1 border-t border-gray-500 pt-2">
                  <p className="text-sm font-semibold text-slate-800">Authorized Signature</p>
                  <p className="mt-0.5 text-xs text-slate-600">
                    {printCompanySettings.name || "Company"}
                  </p>
                </div>
              </div>
            )}
            {(printCompanySettings.bankName ||
              printCompanySettings.accountNumber ||
              printCompanySettings.ifscCode ||
              printCompanySettings.branchName ||
              upiPaymentUri) && (
              <section className="mt-8 h-auto rounded-lg border border-gray-200 bg-gray-50 p-4 text-sm print:mt-3 print:break-inside-avoid print:p-2">
                <div className="flex items-start justify-between gap-6">
                  {(printCompanySettings.bankName ||
                    printCompanySettings.accountNumber ||
                    printCompanySettings.ifscCode ||
                    printCompanySettings.branchName) && (
                    <div className="min-w-0 flex-1">
                      <h2 className="font-bold">Bank Details</h2>
                      <div className="mt-2 grid grid-cols-2 gap-x-6 gap-y-1">
                        {printCompanySettings.bankName && (
                          <p>
                            <span className="font-semibold">Bank Name: </span>
                            {printCompanySettings.bankName}
                          </p>
                        )}
                        {printCompanySettings.accountNumber && (
                          <p>
                            <span className="font-semibold">Account No: </span>
                            {printCompanySettings.accountNumber}
                          </p>
                        )}
                        {printCompanySettings.ifscCode && (
                          <p>
                            <span className="font-semibold">IFSC Code: </span>
                            {printCompanySettings.ifscCode}
                          </p>
                        )}
                        {printCompanySettings.branchName && (
                          <p>
                            <span className="font-semibold">Branch: </span>
                            {printCompanySettings.branchName}
                          </p>
                        )}
                      </div>
                    </div>
                  )}
                  {upiPaymentUri && (
                    <div className="flex shrink-0 flex-col items-center text-center">
                      <QRCode
                        value={upiPaymentUri}
                        size={100}
                        bgColor="#ffffff"
                        fgColor="#000000"
                        aria-label="UPI payment QR code"
                      />
                      <p className="mt-1 text-[10px]">Scan to Pay via UPI</p>
                    </div>
                  )}
                </div>
              </section>
            )}
            <section className="mt-4 h-auto w-full max-w-full break-words whitespace-normal rounded-md border border-gray-300 bg-gray-50 p-4 print:mt-2 print:break-inside-auto print:p-2">
              <h2 className="font-semibold text-slate-900">Terms &amp; Conditions:</h2>
              <ol className="mt-2 list-decimal space-y-1 break-words whitespace-normal pl-5 text-sm text-gray-700 print:mt-1 print:space-y-0">
                <li>Payment due within 7 days of the invoice date.</li>
                <li>Interest @ 18% p.a. will be charged if payment is delayed.</li>
                <li>Subject to Panchkula jurisdiction only.</li>
              </ol>
            </section>
            <p className="mt-10 border-t border-gray-200 pt-4 text-center text-sm text-slate-500 print:mt-2 print:pt-2">
              Thank you for your business!
            </p>
            <p className="mt-1 text-center text-xs text-gray-500">
              This is a computer generated invoice.
            </p>
          </footer>
        </div>
      )}
    </main>
  );
}
