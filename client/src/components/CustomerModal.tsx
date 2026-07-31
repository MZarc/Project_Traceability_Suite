import { useState, useEffect, useMemo } from 'react';
import axios from 'axios';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  Building2, Search, Plus, User as UserIcon, Phone, Mail, Edit2, Trash2, 
  CheckCircle2, AlertCircle, X, Download, MapPin, 
  CreditCard, RefreshCw, ChevronLeft, ChevronRight, ChevronDown, FolderKanban
} from 'lucide-react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const API_URL = 'http://localhost:5000/api';

export interface CustomerContact {
  id?: number | string;
  contact_person: string;
  contact_no: string;
  email: string;
  note?: string;
}

export interface Customer {
  id: number | string;
  name: string;
  address?: string;
  gst_no?: string;
  remarks?: string;
  contacts?: CustomerContact[];
  created_at?: string;
}

interface CustomerModalProps {
  isOpen: boolean;
  onClose: () => void;
  projects?: any[];
  initialSearch?: string;
}

const MOTION_EASE: [number, number, number, number] = [0.22, 1, 0.36, 1];

export default function CustomerModal({ isOpen, onClose, projects = [], initialSearch = '' }: CustomerModalProps) {
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  
  // Search State
  const [searchQuery, setSearchQuery] = useState(initialSearch);

  // Sync initialSearch if prop changes
  useEffect(() => {
    setSearchQuery(initialSearch || '');
  }, [initialSearch]);

  useEffect(() => {
    if (isOpen) {
      loadCustomers();
    }
  }, [isOpen, projects.length]);

  // Automatically open Customer Detail Window when initialSearch is provided (e.g. clicked customer name from table)
  useEffect(() => {
    if (isOpen && initialSearch && customers.length > 0) {
      const q = initialSearch.trim().toLowerCase();
      const match = customers.find(c => c.name.trim().toLowerCase() === q)
                 || customers.find(c => c.name.trim().toLowerCase().includes(q));
      if (match) {
        setSelectedCustomer(match);
        setShowViewModal(true);
      }
    }
  }, [isOpen, initialSearch, customers]);

  // Complete state reset when Customer Modal is closed
  useEffect(() => {
    if (!isOpen) {
      setSearchQuery('');
      setShowViewModal(false);
      setShowAddEditModal(false);
      setShowDeleteAlert(false);
      setSelectedCustomer(null);
      setCurrentPage(1);
    }
  }, [isOpen]);

  // Pagination State
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(9); // 9, 18, 36, 100
  const [isPageSizeOpen, setIsPageSizeOpen] = useState(false);

  // Sub Modals
  const [showAddEditModal, setShowAddEditModal] = useState(false);
  const [showViewModal, setShowViewModal] = useState(false);
  const [showDeleteAlert, setShowDeleteAlert] = useState(false);
  const [selectedCustomer, setSelectedCustomer] = useState<Customer | null>(null);

  // Form State (Credibility removed)
  const [formData, setFormData] = useState({
    name: '',
    address: '',
    gst_no: '',
    remarks: '',
    contacts: [{ contact_person: '', contact_no: '', email: '' }] as CustomerContact[]
  });

  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' } | null>(null);

  const showToast = (message: string, type: 'success' | 'error' = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  };

  // Auto-populate customer names from workbook rows
  const loadCustomers = async () => {
    try {
      setLoading(true);
      const res = await axios.get(`${API_URL}/customers`);
      let fetched: Customer[] = res.data || [];

      if (projects && projects.length > 0) {
        const existingNames = new Set(fetched.map(c => c.name.trim().toLowerCase()));
        const newNames = new Set<string>();

        projects.forEach(p => {
          if (p.customer && typeof p.customer === 'string') {
            const trimmed = p.customer.trim();
            if (trimmed && trimmed.toLowerCase() !== 'n/a' && trimmed.toLowerCase() !== 'na' && !existingNames.has(trimmed.toLowerCase())) {
              newNames.add(trimmed);
            }
          }
        });

        if (newNames.size > 0) {
          for (const name of Array.from(newNames)) {
            try {
              await axios.post(`${API_URL}/customers`, {
                name,
                address: '',
                gst_no: '',
                remarks: 'Auto-populated from project traceability workbook',
                contacts: []
              });
            } catch (e) {
              // ignore duplicate error
            }
          }
          const updatedRes = await axios.get(`${API_URL}/customers`);
          fetched = updatedRes.data || [];
        }
      }

      setCustomers(fetched);
    } catch (err: any) {
      console.error('Failed to load customers:', err);
      showToast('Failed to load customer database', 'error');
    } finally {
      setLoading(false);
    }
  };

  // Reset pagination on search or page size change
  useEffect(() => {
    setCurrentPage(1);
  }, [searchQuery, pageSize]);

  // Filtered customer list - Supports searching by Customer Name, GST, Address, Contacts AND Project Name / SR No!
  const filteredCustomers = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return customers;

    return customers.filter(c => {
      const cName = (c.name || '').toLowerCase();
      const cAddress = (c.address || '').toLowerCase();
      const cGst = (c.gst_no || '').toLowerCase();

      const directMatch =
        cName.includes(q) ||
        cAddress.includes(q) ||
        cGst.includes(q) ||
        (c.contacts && c.contacts.some(ct =>
          (ct.contact_person && ct.contact_person.toLowerCase().includes(q)) ||
          (ct.contact_no && ct.contact_no.includes(q)) ||
          (ct.email && ct.email.toLowerCase().includes(q))
        ));

      if (directMatch) return true;

      // Project Name / SR No search support
      return projects.some(p => {
        const pCust = (p.customer || '').trim().toLowerCase();
        if (pCust !== cName.trim()) return false;
        const pName = (p.project_name || '').toLowerCase();
        const pSr = String(p.sr_no || '').toLowerCase();
        return pName.includes(q) || pSr.includes(q);
      });
    });
  }, [customers, searchQuery, projects]);

  // Robust Pagination Calculation
  const totalPages = Math.ceil(filteredCustomers.length / pageSize) || 1;

  const paginatedCustomers = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredCustomers.slice(start, start + pageSize);
  }, [filteredCustomers, currentPage, pageSize]);

  // Connected Projects list for selected customer in View drawer
  const connectedProjects = useMemo(() => {
    if (!selectedCustomer) return [];
    return projects.filter(p => p.customer && p.customer.trim().toLowerCase() === selectedCustomer.name.trim().toLowerCase());
  }, [projects, selectedCustomer]);

  const handleAddContact = () => {
    setFormData(prev => ({
      ...prev,
      contacts: [...prev.contacts, { contact_person: '', contact_no: '', email: '', note: '' }]
    }));
  };

  const handleRemoveContact = (index: number) => {
    setFormData(prev => ({
      ...prev,
      contacts: prev.contacts.filter((_, i) => i !== index)
    }));
  };

  const handleContactChange = (index: number, field: keyof CustomerContact, value: string) => {
    const newContacts = [...formData.contacts];
    newContacts[index] = { ...newContacts[index], [field]: value };
    setFormData({ ...formData, contacts: newContacts });
  };

  const resetForm = () => {
    setFormData({
      name: '',
      address: '',
      gst_no: '',
      remarks: '',
      contacts: [{ contact_person: '', contact_no: '', email: '', note: '' }]
    });
    setSelectedCustomer(null);
  };

  const handleOpenAdd = () => {
    resetForm();
    setShowAddEditModal(true);
  };

  const handleOpenEdit = (customer: Customer) => {
    setSelectedCustomer(customer);
    setFormData({
      name: customer.name,
      address: customer.address || '',
      gst_no: customer.gst_no || '',
      remarks: customer.remarks || '',
      contacts: customer.contacts && customer.contacts.length > 0 
        ? customer.contacts.map(c => ({
            contact_person: c.contact_person || '',
            contact_no: c.contact_no || '',
            email: c.email || '',
            note: c.note || ''
          }))
        : [{ contact_person: '', contact_no: '', email: '', note: '' }]
    });
    setShowViewModal(false);
    setShowAddEditModal(true);
  };

  const handleOpenView = (customer: Customer) => {
    setSelectedCustomer(customer);
    setShowViewModal(true);
  };

  const handleSave = async () => {
    if (!formData.name.trim()) {
      showToast('Customer name is required', 'error');
      return;
    }

    try {
      setIsSaving(true);
      const cleanedContacts = formData.contacts.filter(c => c.contact_person.trim() !== '');

      const payload = {
        ...formData,
        contacts: cleanedContacts
      };

      if (selectedCustomer) {
        await axios.put(`${API_URL}/customers/${selectedCustomer.id}`, payload);
        showToast('Customer updated successfully');
      } else {
        await axios.post(`${API_URL}/customers`, payload);
        showToast('Customer added successfully');
      }

      setShowAddEditModal(false);
      resetForm();
      loadCustomers();
    } catch (err: any) {
      console.error(err);
      showToast('Failed to save customer', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!selectedCustomer) return;

    try {
      setIsSaving(true);
      await axios.delete(`${API_URL}/customers/${selectedCustomer.id}`);
      showToast('Customer deleted successfully');
      setShowDeleteAlert(false);
      setShowViewModal(false);
      setSelectedCustomer(null);
      loadCustomers();
    } catch (err: any) {
      console.error(err);
      showToast('Failed to delete customer', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const exportToExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Customers Database');

      sheet.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Customer Name', key: 'name', width: 30 },
        { header: 'GST No', key: 'gst_no', width: 20 },
        { header: 'Address', key: 'address', width: 40 },
        { header: 'Remarks / Notes', key: 'remarks', width: 30 },
        { header: 'Contact Person', key: 'contact_person', width: 25 },
        { header: 'Contact Phone', key: 'contact_no', width: 18 },
        { header: 'Contact Email', key: 'email', width: 25 },
        { header: 'Contact Note / Designation', key: 'note', width: 28 }
      ];

      filteredCustomers.forEach(c => {
        if (c.contacts && c.contacts.length > 0) {
          c.contacts.forEach(ct => {
            sheet.addRow({
              id: c.id,
              name: c.name,
              gst_no: c.gst_no || '',
              address: c.address || '',
              remarks: c.remarks || '',
              contact_person: ct.contact_person || '',
              contact_no: ct.contact_no || '',
              email: ct.email || '',
              note: ct.note || ''
            });
          });
        } else {
          sheet.addRow({
            id: c.id,
            name: c.name,
            gst_no: c.gst_no || '',
            address: c.address || '',
            remarks: c.remarks || '',
            contact_person: '',
            contact_no: '',
            email: '',
            note: ''
          });
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `Customers_Database_${new Date().toISOString().split('T')[0]}.xlsx`);
      showToast('Exported customer database to Excel');
    } catch (err) {
      console.error(err);
      showToast('Export failed', 'error');
    }
  };

  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div className="dialog-backdrop fixed inset-0 z-[4000] flex items-center justify-center bg-slate-950/75 p-3 sm:p-4 backdrop-blur-md overflow-y-auto">
        <motion.div
          initial={{ scale: 0.96, opacity: 0, y: 12 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          exit={{ scale: 0.96, opacity: 0, y: 8 }}
          transition={{ duration: 0.16, ease: MOTION_EASE }}
          className="relative w-full max-w-4xl my-auto overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-2xl text-slate-900 flex flex-col max-h-[86vh]"
        >
          {/* Toast Notification Banner */}
          {toast && (
            <div className={`absolute top-2.5 left-1/2 -translate-x-1/2 z-[5000] flex items-center gap-2 rounded-lg px-3 py-1.5 text-xs font-bold shadow-xl border ${
              toast.type === 'success' ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-rose-50 text-rose-800 border-rose-200'
            }`}>
              {toast.type === 'success' ? <CheckCircle2 className="h-3.5 w-3.5 shrink-0 text-emerald-600" /> : <AlertCircle className="h-3.5 w-3.5 shrink-0 text-rose-600" />}
              <span>{toast.message}</span>
            </div>
          )}

          {/* HEADER */}
          <div className="border-b border-slate-100 bg-white px-5 py-3.5 shrink-0 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-sky-500 text-white shadow-md shadow-sky-500/20">
                <Building2 className="h-4.5 w-4.5" />
              </div>
              <div>
                <h2 className="text-base font-black text-slate-900 tracking-tight flex items-center gap-2">
                  <span>Customer Database</span>
                  <span className="rounded-md bg-slate-100 border border-slate-200 px-2 py-0.5 text-[10px] font-bold text-slate-600">
                    {filteredCustomers.length} Records
                  </span>
                </h2>
                <p className="text-[11px] font-medium text-slate-500">Directory of clients, GST records, contacts & linked projects.</p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={exportToExcel}
                className="flex items-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1 text-xs font-bold text-emerald-700 hover:bg-emerald-100 transition-all cursor-pointer"
                title="Export to Excel"
              >
                <Download className="h-3.5 w-3.5" />
                <span>Export</span>
              </button>

              <button
                type="button"
                onClick={handleOpenAdd}
                className="flex items-center gap-1.5 rounded-lg bg-slate-900 px-3 py-1 text-xs font-black text-white shadow-sm hover:bg-slate-800 transition-all cursor-pointer"
              >
                <Plus className="h-3.5 w-3.5" />
                <span>Add Customer</span>
              </button>

              <button
                type="button"
                onClick={onClose}
                className="rounded-lg border border-slate-200 p-1 text-slate-400 hover:bg-slate-50 hover:text-slate-700 transition-all cursor-pointer"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
          </div>

          {/* SEARCH BAR (Compact & Sleek) */}
          <div className="border-b border-slate-100 bg-slate-50/70 px-5 py-2.5 shrink-0 flex items-center gap-3">
            <div className="relative flex-1">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-slate-400" />
              <input
                type="text"
                placeholder="Search customers, GST, contacts, or project names (e.g. XJ1234)..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full rounded-lg border border-slate-200 bg-white pl-9 pr-8 py-1.5 text-xs font-semibold text-slate-800 placeholder-slate-400 focus:border-sky-500 focus:outline-none shadow-sm"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              )}
            </div>
          </div>

          {/* CUSTOMERS GRID */}
          <div className="flex-1 overflow-y-auto p-4 bg-slate-50/30 dialog-scroll">
            {loading ? (
              <div className="flex h-40 items-center justify-center text-slate-400 gap-2">
                <RefreshCw className="h-4 w-4 animate-spin text-sky-600" />
                <span className="text-xs font-bold">Loading customer directory...</span>
              </div>
            ) : filteredCustomers.length === 0 ? (
              <div className="flex h-40 flex-col items-center justify-center text-slate-400 space-y-1">
                <Building2 className="h-7 w-7 text-slate-300" />
                <p className="text-xs font-bold">No customer records match "{searchQuery}".</p>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                {paginatedCustomers.map(cust => {
                  const primaryContact = cust.contacts && cust.contacts[0];
                  // Compute linked projects count
                  const projectCount = projects.filter(p => 
                    p.customer && p.customer.trim().toLowerCase() === cust.name.trim().toLowerCase()
                  ).length;

                  return (
                    <div
                      key={cust.id}
                      onClick={() => handleOpenView(cust)}
                      className="group relative rounded-xl border border-slate-200 bg-white p-3.5 transition-all duration-200 hover:border-sky-300 hover:shadow-md cursor-pointer flex flex-col justify-between"
                    >
                      <div className="space-y-2">
                        <div className="flex items-start justify-between gap-2">
                          <div className="min-w-0 flex-1">
                            <h3 className="text-xs font-black text-slate-900 truncate group-hover:text-sky-600 transition-colors">
                              {cust.name}
                            </h3>
                            {cust.gst_no ? (
                              <p className="text-[10px] font-semibold text-slate-500 flex items-center gap-1 mt-0.5">
                                <CreditCard className="h-3 w-3 text-slate-400 shrink-0" />
                                <span>GST: {cust.gst_no}</span>
                              </p>
                            ) : (
                              <p className="text-[10px] font-medium text-slate-400 italic mt-0.5">No GST No</p>
                            )}
                          </div>

                          {/* Projects Count Badge (Requirement 2) */}
                          <span className="shrink-0 inline-flex items-center gap-1 rounded-md bg-sky-50 border border-sky-200/80 px-2 py-0.5 text-[9px] font-black text-sky-700">
                            <FolderKanban className="h-2.5 w-2.5 text-sky-600" />
                            <span>{projectCount} Project{projectCount === 1 ? '' : 's'}</span>
                          </span>
                        </div>

                        {cust.address && (
                          <p className="text-[10px] font-medium text-slate-600 line-clamp-2 flex items-start gap-1">
                            <MapPin className="h-3 w-3 text-slate-400 shrink-0 mt-0.5" />
                            <span>{cust.address}</span>
                          </p>
                        )}

                        {primaryContact && (
                          <div className="rounded-lg bg-slate-50 p-2 border border-slate-100 space-y-0.5">
                            <p className="text-[10px] font-bold text-slate-800 flex items-center gap-1">
                              <UserIcon className="h-3 w-3 text-sky-600" />
                              <span className="truncate">{primaryContact.contact_person}</span>
                            </p>
                            <div className="flex flex-wrap items-center gap-2 text-[9px] font-semibold text-slate-500">
                              {primaryContact.contact_no && (
                                <span className="flex items-center gap-0.5">
                                  <Phone className="h-2.5 w-2.5 text-slate-400" /> {primaryContact.contact_no}
                                </span>
                              )}
                              {primaryContact.email && (
                                <span className="flex items-center gap-0.5 truncate max-w-[140px]">
                                  <Mail className="h-2.5 w-2.5 text-slate-400" /> {primaryContact.email}
                                </span>
                              )}
                            </div>
                          </div>
                        )}
                      </div>

                      <div className="mt-2.5 pt-2 border-t border-slate-100 flex items-center justify-between text-[10px] font-bold text-slate-400">
                        <span>{cust.contacts ? `${cust.contacts.length} Contact${cust.contacts.length === 1 ? '' : 's'}` : '0 Contacts'}</span>
                        <button
                          type="button"
                          onClick={(e) => { e.stopPropagation(); handleOpenEdit(cust); }}
                          className="p-1 text-slate-400 hover:text-sky-600 transition-colors opacity-0 group-hover:opacity-100"
                          title="Edit Customer"
                        >
                          <Edit2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* FOOTER PAGINATION BAR (Custom UI Dropdown) */}
          <div className="border-t border-slate-100 bg-white px-5 py-2.5 shrink-0 flex items-center justify-between flex-wrap gap-2 text-xs">
            <div className="flex items-center gap-3 text-slate-500 font-semibold">
              <span>
                Showing {filteredCustomers.length === 0 ? 0 : (currentPage - 1) * pageSize + 1} - {Math.min(currentPage * pageSize, filteredCustomers.length)} of {filteredCustomers.length} Customers
              </span>

              {/* Custom Non-Native Dropdown for Page Size */}
              <div className="relative border-l border-slate-200 pl-3">
                <button
                  type="button"
                  onClick={() => setIsPageSizeOpen(!isPageSizeOpen)}
                  className="flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2 py-0.5 text-xs font-bold text-slate-700 hover:bg-slate-100 transition-colors cursor-pointer"
                >
                  <span>{pageSize} / page</span>
                  <ChevronDown className="h-3 w-3 text-slate-400" />
                </button>

                {isPageSizeOpen && (
                  <div className="absolute bottom-full mb-1 left-3 z-50 rounded-lg border border-slate-200 bg-white shadow-lg py-1 w-24">
                    {[9, 18, 36, 100].map(size => (
                      <button
                        key={size}
                        type="button"
                        onClick={() => {
                          setPageSize(size);
                          setIsPageSizeOpen(false);
                        }}
                        className={`w-full px-3 py-1 text-left text-xs font-bold transition-colors cursor-pointer ${
                          pageSize === size ? 'bg-sky-50 text-sky-700' : 'text-slate-700 hover:bg-slate-50'
                        }`}
                      >
                        {size} / page
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </div>

            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-2.5 py-0.5 text-xs font-bold text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:pointer-events-none transition-all cursor-pointer shadow-sm"
              >
                <ChevronLeft className="h-3.5 w-3.5" /> Prev
              </button>

              <span className="text-xs font-extrabold text-slate-800 px-1.5">
                Page {currentPage} of {totalPages}
              </span>

              <button
                type="button"
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage >= totalPages}
                className="flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-2.5 py-0.5 text-xs font-bold text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:pointer-events-none transition-all cursor-pointer shadow-sm"
              >
                Next <ChevronRight className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>
        </motion.div>
      </div>

      {/* VIEW CUSTOMER DRAWER/MODAL */}
      {showViewModal && selectedCustomer && (
        <div className="dialog-backdrop fixed inset-0 z-[5000] flex items-center justify-center bg-slate-950/70 p-4 backdrop-blur-md">
          <div className="relative w-full max-w-lg rounded-2xl border border-slate-200 bg-white p-5 text-slate-900 shadow-2xl space-y-3.5 max-h-[85vh] overflow-y-auto">
            <div className="flex items-start justify-between border-b border-slate-100 pb-2.5">
              <div>
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] font-black uppercase tracking-wider text-sky-600">Customer Details</span>
                  <button
                    type="button"
                    onClick={() => handleOpenEdit(selectedCustomer)}
                    title="Edit Customer"
                    className="h-5 w-5 rounded border border-sky-200 bg-sky-50 text-sky-700 hover:bg-sky-100 flex items-center justify-center transition-all cursor-pointer shadow-sm"
                  >
                    <Edit2 className="h-2.5 w-2.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowDeleteAlert(true)}
                    title="Delete Customer"
                    className="h-5 w-5 rounded border border-rose-200 bg-rose-50 text-rose-700 hover:bg-rose-100 flex items-center justify-center transition-all cursor-pointer shadow-sm"
                  >
                    <Trash2 className="h-2.5 w-2.5" />
                  </button>
                </div>
                <h3 className="text-base font-black text-slate-900 leading-tight mt-0.5">{selectedCustomer.name}</h3>
              </div>

              <button 
                type="button" 
                onClick={() => setShowViewModal(false)} 
                title="Close"
                className="h-6 w-6 rounded-md bg-slate-100 text-slate-400 hover:bg-slate-200 hover:text-slate-700 flex items-center justify-center cursor-pointer transition-colors shrink-0"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            </div>

            {selectedCustomer.gst_no && (
              <div className="rounded-lg border border-slate-100 bg-slate-50 p-2.5 text-xs">
                <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">GST Number</span>
                <p className="font-bold text-slate-900 mt-0.5">{selectedCustomer.gst_no}</p>
              </div>
            )}

            {selectedCustomer.address && (
              <div className="rounded-lg border border-slate-100 bg-slate-50 p-2.5 text-xs space-y-0.5">
                <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Address</span>
                <p className="font-semibold text-slate-800">{selectedCustomer.address}</p>
              </div>
            )}

            {/* Contact Persons List */}
            <div className="space-y-1.5">
              <h4 className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Contact Persons ({selectedCustomer.contacts?.length || 0})</h4>
              {selectedCustomer.contacts && selectedCustomer.contacts.length > 0 ? (
                <div className="space-y-2">
                  {selectedCustomer.contacts.map((c, i) => (
                    <div key={i} className="rounded-xl border border-slate-200/80 bg-slate-50/80 p-3 text-xs space-y-1.5">
                      <div className="flex items-center justify-between">
                        <p className="font-black text-sky-700 text-xs flex items-center gap-1.5">
                          <UserIcon className="h-3.5 w-3.5 text-sky-600 shrink-0" /> {c.contact_person}
                        </p>
                        {c.note && (
                          <span className="px-2 py-0.5 rounded-md bg-sky-100 text-sky-800 text-[10px] font-bold border border-sky-200">
                            {c.note}
                          </span>
                        )}
                      </div>
                      <div className="grid grid-cols-2 gap-2 pt-1 text-[11px] text-slate-600 border-t border-slate-200/50">
                        {c.contact_no && (
                          <div className="flex items-center gap-1.5">
                            <Phone className="h-3 w-3 text-slate-400 shrink-0" />
                            <span className="font-semibold text-slate-800">{c.contact_no}</span>
                          </div>
                        )}
                        {c.email && (
                          <div className="flex items-center gap-1.5 min-w-0">
                            <Mail className="h-3 w-3 text-slate-400 shrink-0" />
                            <span className="font-semibold text-slate-800 truncate">{c.email}</span>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-xs text-slate-400 italic">No contacts registered.</p>
              )}
            </div>

            {/* Connected Projects */}
            <div className="space-y-1.5">
              <h4 className="text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Connected Projects ({connectedProjects.length})
              </h4>
              {connectedProjects.length > 0 ? (
                <div className="max-h-36 overflow-y-auto space-y-1 dialog-scroll pr-1">
                  {connectedProjects.map((p: any) => (
                    <div key={p.id} className="flex items-center justify-between rounded-lg bg-slate-50 p-2 text-xs border border-slate-100">
                      <span className="font-bold text-slate-800">{p.project_name} (SR: {p.sr_no})</span>
                      <span className="text-[10px] font-bold text-sky-600">Qty: {p.qty}</span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-xs text-slate-400 italic">No active projects linked to this customer.</p>
              )}
            </div>

            {selectedCustomer.remarks && (
              <div className="rounded-lg border border-slate-100 bg-slate-50 p-2.5 text-xs space-y-0.5">
                <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Remarks</span>
                <p className="text-slate-700">{selectedCustomer.remarks}</p>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ADD / EDIT FORM MODAL */}
      {showAddEditModal && (
        <div className="dialog-backdrop fixed inset-0 z-[5500] flex items-center justify-center bg-slate-950/70 p-4 backdrop-blur-md overflow-y-auto">
          <div className="relative w-full max-w-lg rounded-2xl border border-slate-200 bg-white p-5 text-slate-900 shadow-2xl space-y-3.5 my-auto max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2.5">
              <h3 className="text-base font-black text-slate-900">
                {selectedCustomer ? 'Edit Customer' : 'Add New Customer'}
              </h3>
              <button type="button" onClick={() => setShowAddEditModal(false)} className="text-slate-400 hover:text-slate-700 p-1">
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                  Customer Name *
                </label>
                <input
                  type="text"
                  placeholder="e.g. MASIBUS, IDEX, SIGNITIVE..."
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="ctrl-input !h-[32px] text-xs font-semibold text-slate-900 w-full"
                />
              </div>

              <div>
                <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                  GST Number
                </label>
                <input
                  type="text"
                  placeholder="24AAAAA0000A1Z5"
                  value={formData.gst_no}
                  onChange={(e) => setFormData({ ...formData, gst_no: e.target.value })}
                  className="ctrl-input !h-[32px] text-xs font-semibold text-slate-900 w-full"
                />
              </div>

              <div>
                <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                  Address
                </label>
                <textarea
                  rows={2}
                  placeholder="Office or plant location..."
                  value={formData.address}
                  onChange={(e) => setFormData({ ...formData, address: e.target.value })}
                  className="ctrl-input !h-auto py-1.5 text-xs font-semibold text-slate-900 w-full"
                />
              </div>

              {/* Dynamic Contacts (Redesigned & Structured) */}
              <div className="space-y-3 pt-3 border-t border-slate-100">
                <div className="flex items-center justify-between">
                  <label className="text-[10px] font-black uppercase tracking-wider text-slate-500">
                    Contact Persons ({formData.contacts.length})
                  </label>
                  <button
                    type="button"
                    onClick={handleAddContact}
                    className="text-[11px] font-bold text-sky-600 hover:text-sky-700 flex items-center gap-1 cursor-pointer transition-colors"
                  >
                    <Plus className="h-3.5 w-3.5" /> Add Contact Person
                  </button>
                </div>

                <div className="space-y-3">
                  {formData.contacts.map((contact, idx) => (
                    <div key={idx} className="rounded-xl border border-slate-200 bg-slate-50/70 p-3 space-y-2.5 relative">
                      <div className="flex items-center justify-between border-b border-slate-200/60 pb-1.5">
                        <span className="text-[10px] font-black uppercase tracking-wider text-slate-500 flex items-center gap-1">
                          <UserIcon className="h-3 w-3 text-sky-600" /> Contact #{idx + 1}
                        </span>
                        {formData.contacts.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleRemoveContact(idx)}
                            className="text-slate-400 hover:text-rose-600 transition-colors p-0.5 cursor-pointer flex items-center gap-1 text-[10px] font-bold"
                            title="Remove contact"
                          >
                            <Trash2 className="h-3 w-3" /> Remove
                          </button>
                        )}
                      </div>

                      <div className="grid grid-cols-2 gap-2.5">
                        <div>
                          <label className="text-[9px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                            Full Name
                          </label>
                          <input
                            type="text"
                            placeholder="e.g. Meet Mistry"
                            value={contact.contact_person}
                            onChange={(e) => handleContactChange(idx, 'contact_person', e.target.value)}
                            className="ctrl-input !h-[32px] text-xs font-semibold text-slate-900 w-full"
                          />
                        </div>

                        <div>
                          <label className="text-[9px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                            Phone Number
                          </label>
                          <input
                            type="text"
                            placeholder="e.g. +91 9876543210"
                            value={contact.contact_no}
                            onChange={(e) => handleContactChange(idx, 'contact_no', e.target.value)}
                            className="ctrl-input !h-[32px] text-xs font-semibold text-slate-900 w-full"
                          />
                        </div>

                        <div>
                          <label className="text-[9px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                            Email Address
                          </label>
                          <input
                            type="email"
                            placeholder="e.g. meetmistryyy@gmail.com"
                            value={contact.email}
                            onChange={(e) => handleContactChange(idx, 'email', e.target.value)}
                            className="ctrl-input !h-[32px] text-xs font-semibold text-slate-900 w-full"
                          />
                        </div>

                        <div>
                          <label className="text-[9px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                            Note / Designation
                          </label>
                          <input
                            type="text"
                            placeholder="e.g. Full-stack developer"
                            value={contact.note || ''}
                            onChange={(e) => handleContactChange(idx, 'note', e.target.value)}
                            className="ctrl-input !h-[32px] text-xs font-semibold text-slate-900 w-full"
                          />
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div>
                <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1 block">
                  Remarks / Notes
                </label>
                <input
                  type="text"
                  placeholder="Special instructions or notes..."
                  value={formData.remarks}
                  onChange={(e) => setFormData({ ...formData, remarks: e.target.value })}
                  className="ctrl-input !h-[32px] text-xs font-semibold text-slate-900 w-full"
                />
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 border-t border-slate-100 pt-2.5">
              <button
                type="button"
                onClick={() => setShowAddEditModal(false)}
                className="rounded-lg border border-slate-200 px-3.5 py-1.5 text-xs font-bold text-slate-600 hover:bg-slate-50 cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSave}
                disabled={isSaving}
                className="rounded-lg bg-slate-900 px-4 py-1.5 text-xs font-black text-white hover:bg-slate-800 transition-all cursor-pointer"
              >
                {isSaving ? 'Saving...' : selectedCustomer ? 'Update Customer' : 'Save Customer'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* DELETE CONFIRMATION ALERT */}
      {showDeleteAlert && selectedCustomer && (
        <div className="dialog-backdrop fixed inset-0 z-[6000] flex items-center justify-center bg-slate-950/70 p-4 backdrop-blur-md">
          <div className="relative w-full max-w-sm rounded-2xl border border-rose-200 bg-white p-5 text-slate-900 shadow-2xl space-y-3 text-center">
            <AlertCircle className="h-8 w-8 text-rose-500 mx-auto" />
            <h3 className="text-base font-black text-slate-900">Delete Customer Record?</h3>
            <p className="text-xs text-slate-500">
              Are you sure you want to delete <strong className="text-slate-900">{selectedCustomer.name}</strong>?
            </p>
            <div className="flex items-center justify-center gap-2 pt-1">
              <button
                type="button"
                onClick={() => setShowDeleteAlert(false)}
                className="rounded-lg border border-slate-200 px-3.5 py-1.5 text-xs font-bold text-slate-600 hover:bg-slate-50 cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleDelete}
                disabled={isSaving}
                className="rounded-lg bg-rose-600 px-4 py-1.5 text-xs font-black text-white hover:bg-rose-500 cursor-pointer"
              >
                {isSaving ? 'Deleting...' : 'Confirm Delete'}
              </button>
            </div>
          </div>
        </div>
      )}
    </AnimatePresence>
  );
}
