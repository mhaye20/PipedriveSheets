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
    
    // Create ticket object
    const ticket = {
      id: ticketId,
      userEmail: userEmail,
      name: ticketData.name,
      email: ticketData.email,
      type: ticketData.type,
      priority: ticketData.priority,
      subject: ticketData.subject,
      description: ticketData.description,
      status: 'open',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      messages: [
        {
          id: Utilities.getUuid(),
          author: ticketData.name,
          authorEmail: ticketData.email,
          message: ticketData.description,
          timestamp: new Date().toISOString(),
          isAdmin: false
        }
      ],
      tags: [],
      assignedTo: null,
      lastUserActivity: new Date().toISOString(),
      lastAdminActivity: null
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
 * Saves a ticket to storage
 * @param {Object} ticket - The ticket object to save
 * @return {boolean} True if saved successfully
 */
SupportTicketService.saveTicket = function(ticket) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    
    // Get existing tickets
    const ticketsJson = scriptProps.getProperty('SUPPORT_TICKETS') || '{}';
    const tickets = JSON.parse(ticketsJson);
    
    // Add/update the ticket
    tickets[ticket.id] = ticket;
    
    // Save back to properties
    scriptProps.setProperty('SUPPORT_TICKETS', JSON.stringify(tickets));
    
    // Also maintain a list of ticket IDs for quick access
    const ticketIds = Object.keys(tickets);
    scriptProps.setProperty('SUPPORT_TICKET_IDS', JSON.stringify(ticketIds));
    
    return true;
  } catch (error) {
    console.error('Error saving ticket:', error);
    return false;
  }
};

/**
 * Gets a ticket by ID
 * @param {string} ticketId - The ticket ID
 * @return {Object|null} The ticket object or null if not found
 */
SupportTicketService.getTicket = function(ticketId) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const ticketsJson = scriptProps.getProperty('SUPPORT_TICKETS') || '{}';
    const tickets = JSON.parse(ticketsJson);
    
    return tickets[ticketId] || null;
  } catch (error) {
    console.error('Error getting ticket:', error);
    return null;
  }
};

/**
 * Gets all tickets for the current user
 * @return {Array} Array of user's tickets
 */
SupportTicketService.getUserTickets = function() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const scriptProps = PropertiesService.getScriptProperties();
    const ticketsJson = scriptProps.getProperty('SUPPORT_TICKETS') || '{}';
    const tickets = JSON.parse(ticketsJson);
    
    const userTickets = [];
    for (const ticketId in tickets) {
      const ticket = tickets[ticketId];
      if (ticket.userEmail === userEmail || ticket.email === userEmail) {
        userTickets.push(ticket);
      }
    }
    
    // Sort by creation date (newest first)
    userTickets.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    
    return userTickets;
  } catch (error) {
    console.error('Error getting user tickets:', error);
    return [];
  }
};

/**
 * Gets all tickets (admin function)
 * @return {Array} Array of all tickets
 */
SupportTicketService.getAllTickets = function() {
  try {
    // Check if user is admin
    if (!this.isAdminUser()) {
      throw new Error('Access denied: Admin privileges required');
    }
    
    const scriptProps = PropertiesService.getScriptProperties();
    const ticketsJson = scriptProps.getProperty('SUPPORT_TICKETS') || '{}';
    const tickets = JSON.parse(ticketsJson);
    
    const allTickets = Object.values(tickets);
    
    // Sort by priority (high first) then by creation date (newest first)
    allTickets.sort((a, b) => {
      const priorityOrder = { 'high': 3, 'medium': 2, 'low': 1 };
      const priorityDiff = priorityOrder[b.priority] - priorityOrder[a.priority];
      if (priorityDiff !== 0) return priorityDiff;
      return new Date(b.createdAt) - new Date(a.createdAt);
    });
    
    return allTickets;
  } catch (error) {
    console.error('Error getting all tickets:', error);
    return [];
  }
};

/**
 * Adds a message to a ticket
 * @param {string} ticketId - The ticket ID
 * @param {string} message - The message content
 * @param {boolean} isAdmin - Whether the message is from admin
 * @return {Object} Result with success status
 */
SupportTicketService.addMessage = function(ticketId, message, isAdmin = false) {
  try {
    const ticket = this.getTicket(ticketId);
    if (!ticket) {
      return { success: false, error: 'Ticket not found' };
    }
    
    const userEmail = Session.getActiveUser().getEmail();
    
    // Verify user can access this ticket
    if (!isAdmin && ticket.userEmail !== userEmail && ticket.email !== userEmail) {
      return { success: false, error: 'Access denied' };
    }
    
    // Create message object
    const messageObj = {
      id: Utilities.getUuid(),
      author: isAdmin ? 'PipedriveSheets Support' : ticket.name,
      authorEmail: isAdmin ? 'support@pipedrivesheets.com' : userEmail,
      message: message,
      timestamp: new Date().toISOString(),
      isAdmin: isAdmin
    };
    
    // Add message to ticket
    ticket.messages.push(messageObj);
    ticket.updatedAt = new Date().toISOString();
    
    if (isAdmin) {
      ticket.lastAdminActivity = new Date().toISOString();
    } else {
      ticket.lastUserActivity = new Date().toISOString();
    }
    
    // Save the updated ticket
    const success = this.saveTicket(ticket);
    
    if (success) {
      // Send notification email to the other party
      if (isAdmin) {
        this.sendUserNotification(ticket, messageObj);
      } else {
        this.sendAdminReplyNotification(ticket, messageObj);
      }
      
      return { success: true };
    } else {
      return { success: false, error: 'Failed to save message' };
    }
  } catch (error) {
    console.error('Error adding message:', error);
    return { success: false, error: error.message };
  }
};

/**
 * Updates ticket status
 * @param {string} ticketId - The ticket ID
 * @param {string} status - New status (open, in_progress, resolved, closed)
 * @return {Object} Result with success status
 */
SupportTicketService.updateTicketStatus = function(ticketId, status) {
  try {
    if (!this.isAdminUser()) {
      return { success: false, error: 'Access denied: Admin privileges required' };
    }
    
    const ticket = this.getTicket(ticketId);
    if (!ticket) {
      return { success: false, error: 'Ticket not found' };
    }
    
    ticket.status = status;
    ticket.updatedAt = new Date().toISOString();
    ticket.lastAdminActivity = new Date().toISOString();
    
    const success = this.saveTicket(ticket);
    
    if (success && status === 'resolved') {
      // Send resolution notification to user
      this.sendResolutionNotification(ticket);
    }
    
    return { success: success };
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
      'support@pipedrivesheets.com',
      'mike@pipedrivesheets.com',
      'mhaye20@gmail.com',
      'connect@mikehaye.com',
      'admin@pipedrivesheets.com'
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
    const subject = `[PipedriveSheets Support] New ${ticket.priority.toUpperCase()} Priority Ticket: ${ticket.subject}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #4285F4;">🎫 New Support Ticket</h2>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <p><strong>Ticket ID:</strong> ${ticket.id}</p>
          <p><strong>From:</strong> ${ticket.name} (${ticket.email})</p>
          <p><strong>Type:</strong> ${ticket.type}</p>
          <p><strong>Priority:</strong> <span style="color: ${ticket.priority === 'high' ? '#dc3545' : ticket.priority === 'medium' ? '#ffc107' : '#28a745'};">${ticket.priority.toUpperCase()}</span></p>
          <p><strong>Subject:</strong> ${ticket.subject}</p>
        </div>
        
        <div style="background: white; border: 1px solid #dee2e6; padding: 20px; border-radius: 6px;">
          <h3>Message:</h3>
          <p style="white-space: pre-wrap;">${ticket.description}</p>
        </div>
        
        <p style="margin-top: 20px;">
          <a href="https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec?mode=admin&ticketId=${ticket.id}" 
             style="background: #4285F4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">
            View in Admin Panel
          </a>
        </p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: 'support@pipedrivesheets.com',
      subject: subject,
      htmlBody: htmlBody,
      name: 'PipedriveSheets Support System'
    });
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
    const subject = `✅ Support Request Received - Ticket #${ticket.id}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #4285F4;">Thanks for contacting PipedriveSheets Support! 🎉</h2>
        
        <p>Hi ${ticket.name},</p>
        
        <p>We've received your support request and will respond within 24 hours. Here are the details:</p>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <p><strong>Ticket ID:</strong> ${ticket.id}</p>
          <p><strong>Subject:</strong> ${ticket.subject}</p>
          <p><strong>Priority:</strong> ${ticket.priority.charAt(0).toUpperCase() + ticket.priority.slice(1)}</p>
          <p><strong>Type:</strong> ${ticket.type}</p>
        </div>
        
        <p>Please keep this ticket ID for your records. You can reply to this email or access your ticket anytime through the PipedriveSheets add-on by going to Help & About → Support → View My Support Tickets.</p>
        
        <p>Thanks for using PipedriveSheets!</p>
        
        <p>Best regards,<br>
        Mike Haye<br>
        PipedriveSheets Support</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: ticket.email,
      subject: subject,
      htmlBody: htmlBody,
      name: 'PipedriveSheets Support',
      replyTo: 'support@pipedrivesheets.com'
    });
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
    const subject = `💬 Reply to your support ticket #${ticket.id}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #4285F4;">New Reply to Your Support Ticket</h2>
        
        <p>Hi ${ticket.name},</p>
        
        <p>We've replied to your support ticket. Here's the response:</p>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <p><strong>Ticket ID:</strong> ${ticket.id}</p>
          <p><strong>Subject:</strong> ${ticket.subject}</p>
        </div>
        
        <div style="background: white; border: 1px solid #dee2e6; padding: 20px; border-radius: 6px;">
          <p><strong>Support Team:</strong></p>
          <p style="white-space: pre-wrap;">${message.message}</p>
        </div>
        
        <p>You can reply to this email or access your ticket through the PipedriveSheets add-on by going to Help & About → Support → View My Support Tickets.</p>
        
        <p>Best regards,<br>
        PipedriveSheets Support</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: ticket.email,
      subject: subject,
      htmlBody: htmlBody,
      name: 'PipedriveSheets Support',
      replyTo: 'support@pipedrivesheets.com'
    });
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
    const subject = `[PipedriveSheets Support] User Reply - Ticket #${ticket.id}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #4285F4;">💬 User Reply</h2>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <p><strong>Ticket ID:</strong> ${ticket.id}</p>
          <p><strong>From:</strong> ${ticket.name} (${ticket.email})</p>
          <p><strong>Subject:</strong> ${ticket.subject}</p>
        </div>
        
        <div style="background: white; border: 1px solid #dee2e6; padding: 20px; border-radius: 6px;">
          <p><strong>New Message:</strong></p>
          <p style="white-space: pre-wrap;">${message.message}</p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: 'support@pipedrivesheets.com',
      subject: subject,
      htmlBody: htmlBody,
      name: 'PipedriveSheets Support System'
    });
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
    const subject = `✅ Support ticket #${ticket.id} has been resolved`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #28a745;">🎉 Ticket Resolved!</h2>
        
        <p>Hi ${ticket.name},</p>
        
        <p>Great news! Your support ticket has been marked as resolved.</p>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <p><strong>Ticket ID:</strong> ${ticket.id}</p>
          <p><strong>Subject:</strong> ${ticket.subject}</p>
          <p><strong>Status:</strong> Resolved ✅</p>
        </div>
        
        <p>If you need further assistance with this issue, simply reply to this email and we'll reopen the ticket.</p>
        
        <p>Thanks for using PipedriveSheets!</p>
        
        <p>Best regards,<br>
        PipedriveSheets Support</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: ticket.email,
      subject: subject,
      htmlBody: htmlBody,
      name: 'PipedriveSheets Support',
      replyTo: 'support@pipedrivesheets.com'
    });
  } catch (error) {
    console.error('Error sending resolution notification:', error);
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