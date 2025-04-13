Update an activity

Updates the properties of an activity.

Request
PATCH/api/v2/activities/{id}
Path parameters
id
integer
required
The ID of the activity

Body parameters
application/json

subject
string
The subject of the activity

type
string
The type of the activity

owner_id
integer
The ID of the user who owns the activity

deal_id
integer
The ID of the deal linked to the activity

lead_id
string
The ID of the lead linked to the activity

person_id
integer
The ID of the person linked to the activity

org_id
integer
The ID of the organization linked to the activity

project_id
integer
The ID of the project linked to the activity

due_date
string
The due date of the activity

due_time
string
The due time of the activity

duration
string
The duration of the activity

busy
boolean
Whether the activity marks the assignee as busy or not in their calendar

done
boolean
Whether the activity is marked as done or not

location
object
Location of the activity

participants
array
The participants of the activity

attendees
array
The attendees of the activity

public_description
string
The public description of the activity

priority
integer
The priority of the activity. Mappable to a specific string using activityFields API.

note
string
The note of the activity.                Update a deal
Copy link
Updates the properties of a deal.


Request
PATCH/api/v2/deals/{id}
Path parameters
id
integer
required
The ID of the deal

Body parameters
application/json

title
string
The title of the deal

owner_id
integer
The ID of the user who owns the deal

person_id
integer
The ID of the person linked to the deal

org_id
integer
The ID of the organization linked to the deal

pipeline_id
integer
The ID of the pipeline associated with the deal

stage_id
integer
The ID of the deal stage

value
number
The value of the deal

currency
string
The currency associated with the deal

add_time
string
The creation date and time of the deal

update_time
string
The last updated date and time of the deal

stage_change_time
string
The last updated date and time of the deal stage

is_deleted
boolean
Whether the deal is deleted or not

status
string
The status of the deal

probability
number
The success probability percentage of the deal

lost_reason
string
The reason for losing the deal. Can only be set if deal status is lost.

visible_to
integer
The visibility of the deal

close_time
string
The date and time of closing the deal. Can only be set if deal status is won or lost.

won_time
string
The date and time of changing the deal status as won. Can only be set if deal status is won.

lost_time
string
The date and time of changing the deal status as lost. Can only be set if deal status is lost.

expected_close_date
string
The expected close date of the deal

Formatdate
label_ids
array
The IDs of labels assigned to the deal.                 Update a deal field
Copy link
Updates a deal field. For more information, see the tutorial for updating custom fields' values.

API v1
Cost
10

Request
PUT/v1/dealFields/{id}
Path parameters
id
integer
required
The ID of the field

Body parameters
application/json

name
string
The name of the field

options
array
When field_type is either set or enum, possible options must be supplied as a JSON-encoded sequential array of objects. All active items must be supplied and already existing items must have their ID supplied. New items only require a label. Example: [{"id":123,"label":"Existing Item"},{"label":"New Item"}]

add_visible_flag
boolean
Whether the field is available in 'add new' modal or not (both in web and mobile app)

Defaulttrue.            Update a lead
Copy link
Updates one or more properties of a lead. Only properties included in the request will be updated. Send null to unset a property (applicable for example for value, person_id or organization_id). If a lead contains custom fields, the fields' values will be included in the response in the same format as with the Deals endpoints. If a custom field's value hasn't been set for the lead, it won't appear in the response. Please note that leads do not have a separate set of custom fields, instead they inherit the custom fields’ structure from deals. See an example given in the updating custom fields’ values tutorial.

API v1
Cost
10

Request
PATCH/v1/leads/{id}
Path parameters
id
string
required
The ID of the lead

Formatuuid
Body parameters
application/json

title
string
The name of the lead

owner_id
integer
The ID of the user which will be the owner of the created lead. If not provided, the user making the request will be used.

label_ids
array
The IDs of the lead labels which will be associated with the lead

person_id
integer
The ID of a person which this lead will be linked to. If the person does not exist yet, it needs to be created first. A lead always has to be linked to a person or organization or both.

organization_id
integer
The ID of an organization which this lead will be linked to. If the organization does not exist yet, it needs to be created first. A lead always has to be linked to a person or organization or both.

is_archived
boolean
A flag indicating whether the lead is archived or not

value
object
The potential value of the lead represented by a JSON object: { "amount": 200, "currency": "EUR" }. Both amount and currency are required.

expected_close_date
string
The date of when the deal which will be created from the lead is expected to be closed. In ISO 8601 format: YYYY-MM-DD.

Formatdate
visible_to
string
The visibility of the lead. If omitted, the visibility will be set to the default visibility setting of this item type for the authorized user. Read more about visibility groups here.

Essential / Advanced plan
Value	Description
1	Owner & followers
3	Entire company
Professional / Enterprise plan
Value	Description
1	Owner only
3	Owner's visibility group
5	Owner's visibility group and sub-groups
7	Entire company
Values

1

3

5

7

was_seen
boolean
A flag indicating whether the lead was seen by someone in the Pipedrive UI

channel
integer
The ID of Marketing channel this lead was created from. Provided value must be one of the channels configured for your company which you can fetch with GET /v1/dealFields.

channel_id
string
The optional ID to further distinguish the Marketing channel.                  Update an organization
Copy link
Updates the properties of a organization.


Request
PATCH/api/v2/organizations/{id}
Path parameters
id
integer
required
The ID of the organization

Body parameters
application/json

name
string
The name of the organization

owner_id
integer
The ID of the user who owns the organization

add_time
string
The creation date and time of the organization

update_time
string
The last updated date and time of the organization

visible_to
integer
The visibility of the organization

label_ids
array
The IDs of labels assigned to the organization.                Update an organization field
Copy link
Updates an organization field. For more information, see the tutorial for updating custom fields' values.

API v1
Cost
10

Request
PUT/v1/organizationFields/{id}
Path parameters
id
integer
required
The ID of the field

Body parameters
application/json

name
string
The name of the field

options
array
When field_type is either set or enum, possible options must be supplied as a JSON-encoded sequential array of objects. All active items must be supplied and already existing items must have their ID supplied. New items only require a label. Example: [{"id":123,"label":"Existing Item"},{"label":"New Item"}]

add_visible_flag
boolean
Whether the field is available in 'add new' modal or not (both in web and mobile app)

Defaulttrue.              Update a person
Copy link
Updates the properties of a person.


Request
PATCH/api/v2/persons/{id}
Path parameters
id
integer
required
The ID of the person

Body parameters
application/json

name
string
The name of the person

owner_id
integer
The ID of the user who owns the person

org_id
integer
The ID of the organization linked to the person

add_time
string
The creation date and time of the person

update_time
string
The last updated date and time of the person

emails
array
The emails of the person

phones
array
The phones of the person

visible_to
integer
The visibility of the person

label_ids
array
The IDs of labels assigned to the person.              Update a person field
Copy link
Updates a person field. For more information, see the tutorial for updating custom fields' values.

API v1
Cost
10

Request
PUT/v1/personFields/{id}
Path parameters
id
integer
required
The ID of the field

Body parameters
application/json

name
string
The name of the field

options
array
When field_type is either set or enum, possible options must be supplied as a JSON-encoded sequential array of objects. All active items must be supplied and already existing items must have their ID supplied. New items only require a label. Example: [{"id":123,"label":"Existing Item"},{"label":"New Item"}]

add_visible_flag
boolean
Whether the field is available in 'add new' modal or not (both in web and mobile app)

Defaulttrue.               Update a product
Copy link
Updates product data.


Request
PATCH/api/v2/products/{id}
Path parameters
id
integer
required
The ID of the product

Body parameters
application/json

name
string
The name of the product. Cannot be an empty string

code
string
The product code

description
string
The product description

unit
string
The unit in which this product is sold

tax
number
The tax percentage

Default0
category
number
The category of the product

owner_id
integer
The ID of the user who will be marked as the owner of this product. When omitted, the authorized user ID will be used

is_linkable
boolean
Whether this product can be added to a deal or not

Defaulttrue
visible_to
number
The visibility of the product. If omitted, the visibility will be set to the default visibility setting of this item type for the authorized user. Read more about visibility groups here.

Essential / Advanced plan
Value	Description
1	Owner & followers
3	Entire company
Professional / Enterprise plan
Value	Description
1	Owner only
3	Owner's visibility group
5	Owner's visibility group and sub-groups
7	Entire company
Values

1

3

5

7

prices
array
An array of objects, each containing: currency (string), price (number), cost (number, optional), direct_cost (number, optional). Note that there can only be one price per product per currency. When prices is omitted altogether, a default price of 0 and the user's default currency will be assigned.

billing_frequency
string
Only available in Advanced and above plans

How often a customer is billed for access to a service or product

Values

one-time

annually

semi-annually

quarterly

monthly

weekly

billing_frequency_cycles
integer
Only available in Advanced and above plans

The number of times the billing frequency repeats for a product in a deal

When billing_frequency is set to one-time, this field must be null

When billing_frequency is set to weekly, this field cannot be null

For all the other values of billing_frequency, null represents a product billed indefinitely

Must be a positive integer less or equal to 208.             Update a product field
Copy link
Updates a product field. For more information, see the tutorial for updating custom fields' values.

API v1
Cost
10

Request
PUT/v1/productFields/{id}
Path parameters
id
integer
required
The ID of the product field

Body parameters
application/json

name
string
The name of the field

options
array
When field_type is either set or enum, possible options on update must be supplied as an array of objects each containing id and label, for example: [{"id":1, "label":"red"},{"id":2, "label":"blue"},{"id":3, "label":"lilac"}].               