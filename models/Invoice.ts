import mongoose, { Schema, type InferSchemaType } from "mongoose";

const consignmentSchema = new Schema(
  {
    docketNo: { type: String, required: true },
    date: { type: String, required: true },
    destination: { type: String, required: true },
    courier: { type: String, required: true },
    mode: { type: String, required: true },
    weight: { type: Number, required: true },
    pieces: { type: Number, required: true },
    amount: { type: Number, required: true },
    taxableValue: { type: Number },
    cgst: { type: Number },
    sgst: { type: Number },
    igst: { type: Number },
  },
  { _id: false },
);

const customerSnapshotSchema = new Schema(
  {
    name: { type: String, required: true },
    phone: { type: String, required: true },
    email: { type: String, default: "" },
    gstin: { type: String, default: "" },
    address: { type: String, required: true },
    state: { type: String, required: true },
    contactPerson: { type: String, default: "" },
    pan: { type: String, default: "" },
    city: { type: String, default: "" },
    stateCode: { type: String, default: "" },
    pincode: { type: String, default: "" },
  },
  { _id: false },
);

const invoiceSchema = new Schema(
  {
    invoiceNo: { type: String, required: true, trim: true },
    invoiceDate: { type: String, required: true },
    dueDate: { type: String, required: true },
    customerName: { type: String, required: true },
    customerState: { type: String, required: true },
    invoiceType: { type: String, enum: ["GST", "Non-GST"], required: true },
    gstType: { type: String, enum: ["IGST", "CGST+SGST"], required: true },
    taxableValue: { type: Number, required: true },
    cgst: { type: Number, required: true },
    sgst: { type: Number, required: true },
    igst: { type: Number, required: true },
    totalAmount: { type: Number, required: true },
    status: { type: String, required: true },
    paymentStatus: { type: String, required: true },
    consignments: { type: [consignmentSchema], required: true, default: [] },
    customer: { type: customerSnapshotSchema, required: true },
    terms: { type: String, required: true, default: "" },
    gstInclusive: { type: Boolean, default: true },
    createdBy: { type: String, trim: true },
  },
  { timestamps: true },
);

export type InvoiceRecord = InferSchemaType<typeof invoiceSchema>;

const Invoice = mongoose.models.Invoice || mongoose.model("Invoice", invoiceSchema);

export default Invoice;
