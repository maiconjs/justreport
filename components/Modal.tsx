import React from 'react';

interface ModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  children: React.ReactNode;
  footer?: React.ReactNode;
}

export const Modal: React.FC<ModalProps> = ({ isOpen, onClose, title, children, footer }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50 backdrop-blur-sm">
      <div className="bg-white rounded-lg shadow-xl w-[500px] max-w-[95vw] flex flex-col max-h-[90vh] animate-fade-in-up">
        <div className="p-4 border-b border-gray-100 flex justify-between items-center">
          <h3 className="text-lg font-bold text-gray-800">{title}</h3>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600">
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"/></svg>
          </button>
        </div>
        <div className="p-4 overflow-y-auto flex-grow custom-scrollbar">
          {children}
        </div>
        {footer && (
          <div className="p-4 border-t border-gray-100 bg-gray-50 rounded-b-lg flex justify-end gap-2">
            {footer}
          </div>
        )}
      </div>
    </div>
  );
};
