import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from django.shortcuts import render
from django.core.mail import EmailMessage
from django.conf import settings
from .forms import UploadFileForm

def convert_excel_to_image_full(file):
    # Read the Excel file
    df = pd.read_excel(file)

    # Check if there is data in the file
    if df.empty:
        raise ValueError('The Excel file is empty.')

    # Get all column names dynamically
    columns = df.columns.tolist()
    results = df.values.tolist()

    # Calculate figure size dynamically
    num_rows = len(results)
    fig_height = max(4, num_rows * 0.3)  # Minimum height of 4 inches
    fig_width = len(columns) * 1.5  # Adjust width based on number of columns

    # Create the figure and axis
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis('off')

    # Create the table
    table = ax.table(cellText=results,
                     colLabels=columns,
                     cellLoc='center',
                     loc='center',
                     bbox=[0, 0, 1, 1])

    # Style the table
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.2)

    # Style column headers
    for (i, label) in enumerate(columns):
        cell = table[0, i]
        cell.set_text_props(weight='bold', color='white')
        cell.set_facecolor('yellow')  # Set header background color to yellow

    # Adjust column widths based on content length
    for i, col in enumerate(columns):
        table.auto_set_column_width([i])

    # Save the image to a BytesIO object
    image_stream = BytesIO()
    plt.savefig(image_stream, format='jpeg', bbox_inches='tight', pad_inches=0.1)
    image_stream.seek(0)

    return image_stream

def convert_excel_to_image_specific(file):
    # Read the Excel file
    df = pd.read_excel(file)

    # Ensure the necessary columns are in the Excel file
    required_columns = ['Cust State', 'Cust Pin', 'DPD']
    if not all(col in df.columns for col in required_columns):
        missing_columns = [col for col in required_columns if col not in df.columns]
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    # Filter the data
    pin_counts = df['Cust Pin'].value_counts()
    repeated_pins = pin_counts[pin_counts >= 2].index
    filtered_df = df[df['Cust Pin'].isin(repeated_pins)]

    # Prepare the results
    results = []
    for pin in repeated_pins:
        pin_data = filtered_df[filtered_df['Cust Pin'] == pin]
        state = pin_data['Cust State'].values[0]
        count = len(pin_data)
        results.append([state, pin, count])
    
    columns = ['Cust State', 'Cust Pin', 'Count']

    # Create the figure and axis
    fig, ax = plt.subplots(figsize=(10, len(results) * 0.5))  # Adjust size based on number of rows
    ax.axis('off')

    # Create the table
    table = ax.table(cellText=results,
                     colLabels=columns,
                     cellLoc='center',
                     loc='center',
                     bbox=[0, 0, 1, 1])

    # Style the table
    table.auto_set_font_size(False)
    table.set_fontsize(12)
    table.scale(1.2, 1.2)
    
    # Style column headers
    for (i, label) in enumerate(columns):
        cell = table[0, i]
        cell.set_text_props(weight='bold', color='white')
        cell.set_facecolor('yellow')  # Set header background color to yellow

    # Set column widths for better readability
    for i in range(len(columns)):
        table.auto_set_column_width([i])
    
    # Save the image to a BytesIO object
    image_stream = BytesIO()
    plt.savefig(image_stream, format='jpeg', bbox_inches='tight', pad_inches=0.1)
    image_stream.seek(0)

    return image_stream

def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            email_to = form.cleaned_data.get('email_to')
            user_name = form.cleaned_data.get('user_name')
            email_body = form.cleaned_data.get('email_body', 'Here is the image generated from the Excel file.')
            image_type = form.cleaned_data.get('image_type', 'full')  # 'full' or 'specific'

            # Construct email subject
            email_subject = f'Python Assignment - {user_name}'

            try:
                # Choose which function to use based on image_type
                if image_type == 'specific':
                    image_stream = convert_excel_to_image_specific(file)
                else:
                    image_stream = convert_excel_to_image_full(file)

                # Send the image via email
                email = EmailMessage(
                    subject=email_subject,
                    body=email_body,
                    from_email=settings.EMAIL_HOST_USER,
                    to=[email_to],
                )
                image_stream.seek(0)  # Ensure the stream is at the beginning
                email.attach('table_image.jpeg', image_stream.read(), 'image/jpeg')
                email.send()

                return render(request, 'success.html')
            except ValueError as e:
                return render(request, 'upload.html', {'form': form, 'error': str(e)})
            except Exception as e:
                return render(request, 'upload.html', {'form': form, 'error': 'An unexpected error occurred. Please try again.'})
    else:
        form = UploadFileForm()
    
    return render(request, 'upload.html', {'form': form})
