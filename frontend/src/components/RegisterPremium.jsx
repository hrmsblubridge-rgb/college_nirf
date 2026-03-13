import React, { useState } from "react";
import { LogOut, User, Mail, Phone, MapPin, GraduationCap, BookOpen, ArrowRight, Building2, Hash } from "lucide-react";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

const INDIA_STATES = [
  "Andhra Pradesh","Arunachal Pradesh","Assam","Bihar","Chhattisgarh","Goa",
  "Gujarat","Haryana","Himachal Pradesh","Jharkhand","Karnataka","Kerala",
  "Madhya Pradesh","Maharashtra","Manipur","Meghalaya","Mizoram","Nagaland",
  "Odisha","Punjab","Rajasthan","Sikkim","Tamil Nadu","Telangana","Tripura",
  "Uttar Pradesh","Uttarakhand","West Bengal","Delhi","Chandigarh",
];

const GRAD_YEARS = Array.from({ length: 13 }, (_, i) => String(2018 + i));

const DEGREES = [
  "B.Tech","B.E.","BCA","B.Sc","B.Com","BA","BBA","MBA","MCA","M.Tech","M.Sc","PhD","Diploma","Other",
];

const COURSES = [
  "Computer Science & Engineering","Information Technology","Electronics & Communication",
  "Mechanical Engineering","Civil Engineering","Electrical Engineering","Business Administration",
  "Commerce","Mathematics","Physics","Biology","Arts & Humanities","Law","Medicine","Other",
];

const FieldLabel = ({ children, required }) => (
  <label className="block mb-2 text-xs font-semibold text-gray-500 uppercase tracking-widest">
    {children}
    {required && <span className="text-[#1A73E8] ml-1">*</span>}
  </label>
);

const PremiumInput = ({ icon: Icon, ...props }) => (
  <div className="relative group">
    {Icon && (
      <div className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400 group-focus-within:text-[#1A73E8] transition-colors duration-200">
        <Icon size={16} />
      </div>
    )}
    <input
      className={`w-full h-12 bg-gray-50 border border-gray-200 rounded-xl text-sm text-gray-800
        placeholder:text-gray-400 focus:outline-none focus:border-[#1A73E8] focus:ring-2
        focus:ring-[#1A73E8]/15 focus:bg-white transition-all duration-200
        hover:border-gray-300 hover:bg-white
        ${Icon ? "pl-10 pr-4" : "px-4"}`}
      {...props}
    />
  </div>
);

const PremiumSelect = ({ placeholder, options, value, onChange, testId }) => (
  <Select value={value} onValueChange={onChange}>
    <SelectTrigger
      data-testid={testId}
      className="h-12 bg-gray-50 border-gray-200 rounded-xl text-sm text-gray-800
        focus:border-[#1A73E8] focus:ring-2 focus:ring-[#1A73E8]/15 focus:bg-white
        hover:border-gray-300 hover:bg-white transition-all duration-200 data-[placeholder]:text-gray-400"
    >
      <SelectValue placeholder={placeholder} />
    </SelectTrigger>
    <SelectContent className="max-h-64 rounded-xl border-gray-100 shadow-xl">
      {options.map((opt) => (
        <SelectItem key={opt} value={opt} className="text-sm rounded-lg">
          {opt}
        </SelectItem>
      ))}
    </SelectContent>
  </Select>
);

const SectionDivider = ({ title }) => (
  <div className="col-span-2 flex items-center gap-3 mt-2 mb-1">
    <div className="h-px flex-1 bg-gray-100" />
    <span className="text-xs font-semibold text-gray-400 uppercase tracking-widest whitespace-nowrap">
      {title}
    </span>
    <div className="h-px flex-1 bg-gray-100" />
  </div>
);

export default function RegisterPremium() {
  const [form, setForm] = useState({
    fullName: "", email: "", phone: "", age: "",
    state: "", city: "", gradYear: "", college: "",
    degree: "", course: "",
  });
  const [confirmed, setConfirmed] = useState(false);
  const [focused, setFocused] = useState(null);

  const set = (key) => (val) => setForm((p) => ({ ...p, [key]: val }));
  const setInput = (key) => (e) => setForm((p) => ({ ...p, [key]: e.target.value }));

  return (
    <div className="min-h-screen flex flex-col" style={{ fontFamily: "'Inter', sans-serif", backgroundColor: "#F5F2E9" }}>

      {/* Google Fonts */}
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Inter:wght@400;500;600&display=swap');`}</style>

      {/* Header */}
      <header
        data-testid="header"
        className="sticky top-0 z-50 flex items-center justify-between px-6 md:px-12 h-16 border-b border-gray-200/60"
        style={{ backgroundColor: "rgba(245,242,233,0.85)", backdropFilter: "blur(12px)" }}
      >
        <div
          data-testid="logo"
          className="text-2xl font-extrabold tracking-tighter select-none"
          style={{ fontFamily: "'Manrope', sans-serif", color: "#111827", letterSpacing: "-0.04em" }}
        >
          BLU<span className="inline-block w-7 h-7 bg-[#1A73E8] text-white text-base font-black rounded-md mx-0.5 leading-7 text-center align-middle">B</span>RIDGE
        </div>
        <button
          data-testid="logout-btn"
          className="flex items-center gap-2 text-sm font-medium text-gray-500 hover:text-[#1A73E8] transition-colors duration-200 group"
        >
          <LogOut size={15} className="group-hover:translate-x-0.5 transition-transform duration-200" />
          Logout
        </button>
      </header>

      {/* Main */}
      <main className="flex-1 flex items-start justify-center px-4 py-12 md:py-16">
        <div className="w-full max-w-3xl">

          {/* Card */}
          <div
            data-testid="registration-card"
            className="bg-white rounded-3xl overflow-hidden"
            style={{ boxShadow: "0 25px 60px -10px rgba(0,0,0,0.12), 0 8px 24px -6px rgba(0,0,0,0.07)" }}
          >
            {/* Blue top accent bar */}
            <div className="h-1.5 w-full" style={{ background: "linear-gradient(90deg, #1A73E8 0%, #4FA3F7 60%, #1A73E8 100%)" }} />

            <div className="px-8 md:px-12 pt-10 pb-12">

              {/* Title */}
              <div className="mb-10 text-center">
                <h1
                  className="text-4xl md:text-5xl font-bold tracking-tight text-gray-900 mb-2"
                  style={{ fontFamily: "'Manrope', sans-serif" }}
                >
                  Registration Form
                </h1>
                <p className="text-sm text-gray-400 font-medium tracking-wide">
                  Fill in your details to get started with BluBridge
                </p>
              </div>

              {/* Form Grid */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-6">

                {/* ── Personal Info Section ── */}
                <SectionDivider title="Personal Information" />

                {/* Full Name */}
                <div data-testid="field-full-name">
                  <FieldLabel required>Full Name:</FieldLabel>
                  <PremiumInput
                    data-testid="input-full-name"
                    icon={User}
                    type="text"
                    placeholder="Enter your full name"
                    value={form.fullName}
                    onChange={setInput("fullName")}
                  />
                </div>

                {/* Email Address */}
                <div data-testid="field-email">
                  <FieldLabel required>Email Address:</FieldLabel>
                  <PremiumInput
                    data-testid="input-email"
                    icon={Mail}
                    type="email"
                    placeholder="Enter your email"
                    value={form.email}
                    onChange={setInput("email")}
                  />
                </div>

                {/* Phone Number */}
                <div data-testid="field-phone">
                  <FieldLabel required>Phone Number:</FieldLabel>
                  <PremiumInput
                    data-testid="input-phone"
                    icon={Phone}
                    type="tel"
                    placeholder="Enter your phone number"
                    value={form.phone}
                    onChange={setInput("phone")}
                  />
                  <p className="mt-1.5 text-xs text-gray-400 flex items-center gap-1">
                    <span className="inline-block w-1 h-1 rounded-full bg-[#1A73E8]" />
                    Note: Active WhatsApp number (Required)
                  </p>
                </div>

                {/* Age */}
                <div data-testid="field-age">
                  <FieldLabel required>Age:</FieldLabel>
                  <PremiumInput
                    data-testid="input-age"
                    icon={Hash}
                    type="number"
                    placeholder="Enter your age"
                    min={15}
                    max={60}
                    value={form.age}
                    onChange={setInput("age")}
                  />
                </div>

                {/* Location Section */}
                <SectionDivider title="Location Details" />

                {/* Current Location (State) */}
                <div data-testid="field-state">
                  <FieldLabel required>Current Location (State):</FieldLabel>
                  <PremiumSelect
                    testId="select-state"
                    placeholder="Select State"
                    options={INDIA_STATES}
                    value={form.state}
                    onChange={set("state")}
                  />
                </div>

                {/* Preferred Location (City) */}
                <div data-testid="field-city">
                  <FieldLabel required>
                    Preferred Location (City)<span className="text-red-500">*</span>:
                  </FieldLabel>
                  <PremiumInput
                    data-testid="input-city"
                    icon={MapPin}
                    type="text"
                    placeholder="Start typing city name..."
                    value={form.city}
                    onChange={setInput("city")}
                  />
                </div>

                {/* Academic Section */}
                <SectionDivider title="Academic Information" />

                {/* Year of Graduation */}
                <div data-testid="field-grad-year">
                  <FieldLabel>Year of Graduation:</FieldLabel>
                  <PremiumSelect
                    testId="select-grad-year"
                    placeholder="Select"
                    options={GRAD_YEARS}
                    value={form.gradYear}
                    onChange={set("gradYear")}
                  />
                </div>

                {/* College */}
                <div data-testid="field-college">
                  <FieldLabel>College:</FieldLabel>
                  <PremiumInput
                    data-testid="input-college"
                    icon={Building2}
                    type="text"
                    placeholder="Enter your college name"
                    value={form.college}
                    onChange={setInput("college")}
                  />
                </div>

                {/* Degree */}
                <div data-testid="field-degree">
                  <FieldLabel>Degree:</FieldLabel>
                  <PremiumSelect
                    testId="select-degree"
                    placeholder="Select a degree"
                    options={DEGREES}
                    value={form.degree}
                    onChange={set("degree")}
                  />
                </div>

                {/* Course */}
                <div data-testid="field-course">
                  <FieldLabel>Course:</FieldLabel>
                  <PremiumSelect
                    testId="select-course"
                    placeholder="Select a course"
                    options={COURSES}
                    value={form.course}
                    onChange={set("course")}
                  />
                </div>

              </div>

              {/* Divider */}
              <div className="my-8 h-px bg-gray-100" />

              {/* Confirmation Checkbox */}
              <div
                data-testid="confirmation-section"
                className={`flex items-start gap-4 p-5 rounded-2xl border transition-all duration-200 ${
                  confirmed
                    ? "bg-blue-50/70 border-blue-200"
                    : "bg-gray-50 border-gray-200 hover:border-gray-300"
                }`}
              >
                <Checkbox
                  data-testid="confirmation-checkbox"
                  id="confirm"
                  checked={confirmed}
                  onCheckedChange={setConfirmed}
                  className="mt-0.5 h-5 w-5 border-2 border-gray-300 data-[state=checked]:bg-[#1A73E8] data-[state=checked]:border-[#1A73E8] rounded-md transition-all duration-200"
                />
                <label
                  htmlFor="confirm"
                  className="text-sm text-gray-600 leading-relaxed cursor-pointer select-none"
                >
                  I hereby confirm that all the information provided above is accurate to the best of my knowledge.
                </label>
              </div>

              {/* PROCEED Button */}
              <div className="mt-8">
                <button
                  data-testid="proceed-btn"
                  disabled={!confirmed}
                  className={`w-full h-14 flex items-center justify-center gap-3 rounded-2xl font-bold text-base tracking-wide text-white transition-all duration-300
                    ${confirmed
                      ? "bg-[#1A73E8] hover:bg-[#1557B0] shadow-lg shadow-blue-500/30 hover:shadow-blue-500/50 hover:-translate-y-0.5 active:translate-y-0 cursor-pointer"
                      : "bg-gray-300 cursor-not-allowed opacity-70"
                    }`}
                  style={{ fontFamily: "'Manrope', sans-serif", letterSpacing: "0.08em" }}
                >
                  PROCEED
                  {confirmed && <ArrowRight size={18} className="transition-transform duration-200 group-hover:translate-x-1" />}
                </button>
              </div>

            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer
        data-testid="footer"
        className="py-6 px-6 text-center"
        style={{ backgroundColor: "#222222" }}
      >
        <p className="text-sm text-gray-400">
          Copyright 2026 &copy; <span className="text-white font-semibold">Blubridge.com</span>
        </p>
      </footer>

    </div>
  );
}
