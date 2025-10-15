// src/components/ui/Modal.jsx
import React, { useEffect } from 'react';
import ReactDOM from 'react-dom';
import './Modal.css';

/**
 * Modal Component
 * @param {boolean} isOpen - Controls visibility
 * @param {function} onClose - Called when modal is closed
 * @param {string} title - Modal title (optional)
 * @param {React.ReactNode} children - Modal content
 * @param {React.ReactNode} footer - Optional footer (e.g., action buttons)
 * @param {string} size - 'sm' | 'md' | 'lg' (default: 'md')
 * @param {boolean} closeOnOverlayClick - Close when clicking outside content (default: true)
 */
const Modal = ({
  isOpen,
  onClose,
  title,
  children,
  footer,
  size = 'md',
  closeOnOverlayClick = true
}) => {

  useEffect(() => {
    const handleEscape = (e) => {
      if (e.key === 'Escape' && isOpen) {
        onClose();
      }
    };

    if (isOpen) {
      document.body.style.overflow = 'hidden';
      window.addEventListener('keydown', handleEscape);
    }

    return () => {
      document.body.style.overflow = '';
      window.removeEventListener('keydown', handleEscape);
    };
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  const handleOverlayClick = (e) => {
    if (e.target === e.currentTarget && closeOnOverlayClick) {
      onClose();
    }
  };

  return ReactDOM.createPortal(
    <div className="modal-overlay" onClick={handleOverlayClick} role="dialog" aria-modal="true">
      <div className={`modal-content modal-content--${size}`}>
        {(title || onClose) && (
          <div className="modal-header">
            {title && <h2 className="modal-title">{title}</h2>}
            <button
              type="button"
              className="modal-close-btn"
              onClick={onClose}
              aria-label="Close modal"
            >
              ✕
            </button>
          </div>
        )}

        <div className="modal-body">{children}</div>
        {footer && <div className="modal-footer">{footer}</div>}
      </div>
    </div>,
    document.body
  );
};

export default Modal;
