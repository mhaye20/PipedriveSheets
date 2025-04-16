class HeaderService {
  static generateHeaders(columns, options = {}) {
    return columns.map(column => {
      // Use custom name if provided
      if (column.name) {
        return column.name;
      }
      
      // Handle nested fields (like email.work)
      const columnKey = column.key || column;
      if (typeof columnKey === 'string' && columnKey.includes('.')) {
        const parts = columnKey.split('.');
        
        // For email.work and phone.work, format as "Email Work" and "Phone Work"
        if ((parts[0] === 'email' || parts[0] === 'phone') && parts.length > 1) {
          // Capitalize both parts: "Email Work" instead of just "Email"
          return parts.map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
        }
        
        // For other nested fields, join with spaces and capitalize
        return parts.map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
      }
      
      // For simple fields, just capitalize
      return columnKey.charAt(0).toUpperCase() + columnKey.slice(1);
    });
  }
  
  // Other methods...
} 