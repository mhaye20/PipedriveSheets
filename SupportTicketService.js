/**
 * Support Ticket Service
 * Handles support ticket creation, management, and chat functionality
 */

var SupportTicketService = SupportTicketService || {};

/**
 * Generates a unique ticket ID
 * @return {string} A formatted ticket ID
 */
SupportTicketService.generateTicketId = function() {
  const timestamp = new Date().getTime().toString().slice(-6);
  const random = Math.random().toString(36).substring(2, 5).toUpperCase();
  return `PSS-${timestamp}-${random}`;
};

/**
 * Creates a new support ticket
 * @param {Object} ticketData - The ticket data from the form
 * @return {Object} Result with success status and ticket ID
 */
SupportTicketService.createTicket = function(ticketData) {
  try {
    const ticketId = this.generateTicketId();
    const userEmail = Session.getActiveUser().getEmail();
    
    // Create ticket object with backend field names
    const ticket = {
      id: ticketId,
      userEmail: userEmail,
      user_name: ticketData.name,  // Changed from 'name'
      contact_email: ticketData.email,  // Changed from 'email'
      type: ticketData.type,
      priority: ticketData.priority,
      subject: ticketData.subject,
      description: ticketData.description
      // Backend will handle: status, timestamps, initial message
    };
    
    // Store the ticket
    const success = this.saveTicket(ticket);
    
    if (success) {
      // Send notification email to admin
      this.sendAdminNotification(ticket);
      
      // Send confirmation email to user
      this.sendUserConfirmation(ticket);
      
      return {
        success: true,
        ticketId: ticketId
      };
    } else {
      return {
        success: false,
        error: 'Failed to save ticket'
      };
    }
  } catch (error) {
    console.error('Error creating ticket:', error);
    return {
      success: false,
      error: error.message
    };
  }
};

/**
 * Saves a ticket to backend storage
 * @param {Object} ticket - The ticket object to save
 * @return {boolean} True if saved successfully
 */
SupportTicketService.saveTicket = function(ticket) {
  try {
    // Get backend URL
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrive-sheets.vercel.app';
    
    // Save to backend
    const response = UrlFetchApp.fetch(`${backendUrl}/api/tickets/create`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(ticket),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to save ticket:', response.getContentText());
      return false;
    }
    
    return true;
  } catch (error) {
    console.error('Error saving ticket:', error);
    return false;
  }
};

/**
 * Gets a ticket by ID from backend
 * @param {string} ticketId - The ticket ID
 * @return {Object|null} The ticket object or null if not found
 */
SupportTicketService.getTicket = function(ticketId) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrive-sheets.vercel.app';
    
    const response = UrlFetchApp.fetch(`${backendUrl}/api/tickets/get`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ ticketId: ticketId }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to get ticket:', response.getContentText());
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    return data.ticket;
  } catch (error) {
    console.error('Error getting ticket:', error);
    return null;
  }
};

/**
 * Gets all tickets for the current user from backend
 * @return {Array} Array of user's tickets
 */
SupportTicketService.getUserTickets = function() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrive-sheets.vercel.app';
    
    const response = UrlFetchApp.fetch(`${backendUrl}/api/tickets/user`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ userEmail: userEmail }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to get user tickets:', response.getContentText());
      return [];
    }
    
    const data = JSON.parse(response.getContentText());
    return data.tickets || [];
  } catch (error) {
    console.error('Error getting user tickets:', error);
    return [];
  }
};

/**
 * Gets all tickets from backend (admin function)
 * @return {Array} Array of all tickets
 */
SupportTicketService.getAllTickets = function() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // Check if user is admin
    if (!this.isAdminUser()) {
      throw new Error('Access denied: Admin privileges required');
    }
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrive-sheets.vercel.app';
    
    const response = UrlFetchApp.fetch(`${backendUrl}/api/tickets/all`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ adminEmail: userEmail }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to get all tickets:', response.getContentText());
      return [];
    }
    
    const data = JSON.parse(response.getContentText());
    return data.tickets || [];
  } catch (error) {
    console.error('Error getting all tickets:', error);
    return [];
  }
};

/**
 * Adds a message to a ticket via backend
 * @param {string} ticketId - The ticket ID
 * @param {string} message - The message content
 * @param {boolean} isAdmin - Whether the message is from admin
 * @return {Object} Result with success status
 */
SupportTicketService.addMessage = function(ticketId, message, isAdmin = false) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrive-sheets.vercel.app';
    
    // Get ticket to verify access and get info for notifications
    const ticket = this.getTicket(ticketId);
    if (!ticket) {
      return { success: false, error: 'Ticket not found' };
    }
    
    // Verify user can access this ticket
    if (!isAdmin && ticket.user_email !== userEmail && ticket.contact_email !== userEmail) {
      return { success: false, error: 'Access denied' };
    }
    
    const author = isAdmin ? 'PipedriveSheets Support' : ticket.user_name;
    const authorEmail = isAdmin ? 'support@pipedrivesheets.com' : userEmail;
    
    // Add message via backend
    const response = UrlFetchApp.fetch(`${backendUrl}/api/tickets/message`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        ticketId: ticketId,
        message: message,
        author: author,
        authorEmail: authorEmail,
        isAdmin: isAdmin
      }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to add message:', response.getContentText());
      return { success: false, error: 'Failed to save message' };
    }
    
    // Send notification email
    const messageObj = {
      message: message,
      author: author,
      authorEmail: authorEmail
    };
    
    if (isAdmin) {
      this.sendUserNotification(ticket, messageObj);
    } else {
      this.sendAdminReplyNotification(ticket, messageObj);
    }
    
    return { success: true };
  } catch (error) {
    console.error('Error adding message:', error);
    return { success: false, error: error.message };
  }
};

/**
 * Updates ticket status via backend
 * @param {string} ticketId - The ticket ID
 * @param {string} status - New status (open, in_progress, resolved, closed)
 * @return {Object} Result with success status
 */
SupportTicketService.updateTicketStatus = function(ticketId, status) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    if (!this.isAdminUser()) {
      return { success: false, error: 'Access denied: Admin privileges required' };
    }
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrive-sheets.vercel.app';
    
    // Update status via backend
    const response = UrlFetchApp.fetch(`${backendUrl}/api/tickets/status`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        ticketId: ticketId,
        status: status,
        adminEmail: userEmail
      }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to update status:', response.getContentText());
      return { success: false, error: 'Failed to update status' };
    }
    
    // If resolved, send notification
    if (status === 'resolved') {
      const ticket = this.getTicket(ticketId);
      if (ticket) {
        this.sendResolutionNotification(ticket);
      }
    }
    
    return { success: true };
  } catch (error) {
    console.error('Error updating ticket status:', error);
    return { success: false, error: error.message };
  }
};

/**
 * Checks if the current user is an admin
 * @return {boolean} True if user is admin
 */
SupportTicketService.isAdminUser = function() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const adminEmails = [
      'mhaye20@gmail.com',
      'support@pipedrivesheets.com',
      'connect@mikehaye.com'
    ];
    
    return adminEmails.includes(userEmail.toLowerCase());
  } catch (error) {
    return false;
  }
};

/**
 * Sends admin notification for new ticket
 * @param {Object} ticket - The ticket object
 */
SupportTicketService.sendAdminNotification = function(ticket) {
  try {
    // Get priority color for template
    const priorityColors = {
      'high': 'DC2626',
      'medium': 'F59E0B', 
      'low': '10B981'
    };
    
    const emailData = {
      to: 'support@pipedrivesheets.com',
      type: 'admin_notification',
      templateId: 3,
      params: {
        TICKET_ID: ticket.id,
        NAME: ticket.user_name,
        EMAIL: ticket.contact_email,
        TYPE: ticket.type,
        PRIORITY: ticket.priority.toUpperCase(),
        PRIORITY_COLOR: priorityColors[ticket.priority] || '6B7280',
        SUBJECT: ticket.subject,
        DESCRIPTION: ticket.description,
        ADMIN_LINK: SpreadsheetApp.getActiveSpreadsheet().getUrl()
      }
    };
    
    this.sendEmail(emailData);
  } catch (error) {
    console.error('Error sending admin notification:', error);
  }
};

/**
 * Sends user confirmation email
 * @param {Object} ticket - The ticket object
 */
SupportTicketService.sendUserConfirmation = function(ticket) {
  try {
    const emailData = {
      to: ticket.contact_email,
      type: 'user_confirmation',
      templateId: 2,
      params: {
        NAME: ticket.user_name,
        TICKET_ID: ticket.id,
        SUBJECT: ticket.subject,
        PRIORITY: ticket.priority.charAt(0).toUpperCase() + ticket.priority.slice(1),
        TYPE: ticket.type
      }
    };
    
    this.sendEmail(emailData);
  } catch (error) {
    console.error('Error sending user confirmation:', error);
  }
};

/**
 * Sends notification to user when admin replies
 * @param {Object} ticket - The ticket object
 * @param {Object} message - The message object
 */
SupportTicketService.sendUserNotification = function(ticket, message) {
  try {
    const emailData = {
      to: ticket.contact_email,
      type: 'admin_reply',
      templateId: 4,
      params: {
        NAME: ticket.user_name,
        TICKET_ID: ticket.id,
        SUBJECT: ticket.subject,
        MESSAGE: message.message
      }
    };
    
    this.sendEmail(emailData);
  } catch (error) {
    console.error('Error sending user notification:', error);
  }
};

/**
 * Sends notification to admin when user replies
 * @param {Object} ticket - The ticket object
 * @param {Object} message - The message object
 */
SupportTicketService.sendAdminReplyNotification = function(ticket, message) {
  try {
    const emailData = {
      to: 'support@pipedrivesheets.com',
      type: 'user_reply',
      templateId: 5,
      params: {
        TICKET_ID: ticket.id,
        NAME: ticket.user_name,
        EMAIL: ticket.contact_email,
        SUBJECT: ticket.subject,
        MESSAGE: message.message,
        ADMIN_LINK: SpreadsheetApp.getActiveSpreadsheet().getUrl()
      }
    };
    
    this.sendEmail(emailData);
  } catch (error) {
    console.error('Error sending admin reply notification:', error);
  }
};

/**
 * Sends resolution notification to user
 * @param {Object} ticket - The ticket object
 */
SupportTicketService.sendResolutionNotification = function(ticket) {
  try {
    const emailData = {
      to: ticket.contact_email,
      type: 'ticket_resolved',
      templateId: 6,
      params: {
        NAME: ticket.user_name,
        TICKET_ID: ticket.id,
        SUBJECT: ticket.subject
      }
    };
    
    this.sendEmail(emailData);
  } catch (error) {
    console.error('Error sending resolution notification:', error);
  }
};

/**
 * Sends email via Brevo backend API
 * @param {Object} emailData - Email data object
 */
SupportTicketService.sendEmail = function(emailData) {
  try {
    // Get backend URL from script properties or use default
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrive-sheets.vercel.app';
    
    // Send email request to backend
    const response = UrlFetchApp.fetch(`${backendUrl}/api/send-support-email`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(emailData),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to send support email:', response.getContentText());
    } else {
      console.log('Support email sent successfully');
    }
  } catch (error) {
    console.error('Error sending support email:', error);
  }
};

/**
 * Gets user information for form pre-population
 * @return {Object} User information
 */
SupportTicketService.getUserInfo = function() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // Try to get name from email (basic attempt)
    let name = '';
    if (userEmail) {
      name = userEmail.split('@')[0].replace(/[._]/g, ' ');
      name = name.split(' ').map(word => 
        word.charAt(0).toUpperCase() + word.slice(1)
      ).join(' ');
    }
    
    return {
      email: userEmail,
      name: name
    };
  } catch (error) {
    return {
      email: '',
      name: ''
    };
  }
};