import mongoose, { Schema, type InferSchemaType } from "mongoose";

const customerSchema = new Schema(
  {
    name: { type: String, required: true, trim: true },
    phone: { type: String, required: true, trim: true },
    email: { type: String, default: "", trim: true },
    gstin: { type: String, default: "", trim: true },
    address: { type: String, required: true, trim: true },
    state: { type: String, required: true, trim: true },
    contactPerson: { type: String, default: "", trim: true },
    pan: { type: String, default: "", trim: true, uppercase: true },
    city: { type: String, default: "", trim: true },
    stateCode: { type: String, default: "", trim: true },
    pincode: { type: String, default: "", trim: true },
    createdBy: { type: String, trim: true },
  },
  { timestamps: true },
);

export type CustomerRecord = InferSchemaType<typeof customerSchema>;

const Customer =
  mongoose.models.Customer || mongoose.model("Customer", customerSchema);

export default Customer;
