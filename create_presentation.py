# create_presentation.py
import win32com.client

def hex_to_rgb(value):
    value = value.lstrip("#")
    length = len(value)
    return tuple(int(value[i:i + length // 3], 16) for i in range(0, length, length // 3))

def rgb_to_int(rgb):
    return int(rgb[0]) | (int(rgb[1]) << 8) | (int(rgb[2]) << 16)

def create_powerpoint_sequence_diagram(participants_info, messages_info, output_path, autonumber, left_start=100, top_start=100, width=800, height=600, messages_per_slide=5):
    # Create PowerPoint application and presentation
    PPTApp = win32com.client.Dispatch("PowerPoint.Application")
    PPTApp.Visible = True
    presentation = PPTApp.Presentations.Add()

    # Calculate the number of slides needed
    num_slides = (len(messages_info) + messages_per_slide - 1) // messages_per_slide

    # Initialize message number
    message_number = 1

    for slide_num in range(num_slides):
        # Add a new slide
        slide = presentation.Slides.Add(slide_num + 1, 12)  # 12 is the layout index for blank slides

        # Calculate spacing based on the number of participants
        participant_width = 100
        participant_height = 50
        total_width = width - (2 * left_start)
        spacing = total_width / (len(participants_info) - 1)

        # Add participants
        left = left_start
        top = top_start
        participant_shapes = []
        for participant_info in participants_info:
            shape = slide.Shapes.AddShape(1, left, top, participant_width, participant_height)
            shape.TextFrame.TextRange.Text = participant_info["text"]
            shape.Line.Visible = False
            if participant_info["color"]:
                shape.Fill.ForeColor.RGB = rgb_to_int(hex_to_rgb(participant_info["color"]))
            participant_shapes.append(shape)

            # Draw vertical line for each participant
            vertical_line = slide.Shapes.AddLine(left + participant_width/2, top + participant_height, left + participant_width/2, top + height)
            vertical_line.Line.ForeColor.RGB = 0x000000

            left += spacing

        # Add messages for the current slide only
        message_top = top + participant_height + 30
        message_spacing = 50
        start_message_idx = slide_num * messages_per_slide
        end_message_idx = min((slide_num + 1) * messages_per_slide, len(messages_info))
        for message_info in messages_info[start_message_idx:end_message_idx]:
            from_participant = participant_shapes[message_info["from"]]
            to_participant = participant_shapes[message_info["to"]]

            # Add message text
            text_left = from_participant.Left + (abs(from_participant.Left - to_participant.Left) / 2) - 50
            # text_shape = slide.Shapes.AddShape(1, text_left, message_top, 100, 20)
            text_shape = slide.Shapes.AddShape(1, text_left, message_top - 20, 100, 20) # Change message_top to message_top - 20
            message_text = message_info["text"]
            if autonumber:
                message_text = f"{message_number}. {message_text}"
                message_number += 1
            text_shape.TextFrame.TextRange.Text = message_text
            text_shape.Line.Visible = False
            text_shape.TextFrame.TextRange.Text = message_text
            text_shape.Line.Visible = False
            text_shape.Fill.ForeColor.RGB = 0x000000
            text_shape.TextFrame.TextRange.Font.Size = 10
            text_shape.TextFrame.AutoSize = 1

            # if autonumber:
            #     # Add circle shape for the message number
            #     number_shape = slide.Shapes.AddShape(9, text_left - 20, message_top, 20, 20) # 9 is the index for oval shape
            #     number_shape.TextFrame.TextRange.Text = str(message_number - 1)
            #     number_shape.Line.Visible = False
            #     number_shape.Fill.ForeColor.RGB = 0xFFFFFF
            #     number_shape.TextFrame.TextRange.Font.Size = 10
            #     number_shape.TextFrame.AutoSize = 1

            # Add horizontal line with arrow
            horizontal_line = slide.Shapes.AddLine(from_participant.Left + participant_width / 2, message_top + 10,
                                                   to_participant.Left + participant_width / 2, message_top + 10)
            horizontal_line.Line.ForeColor.RGB = 0x000000
            horizontal_line.Line.EndArrowheadStyle = 4  # 4 is the arrowhead style
            if message_info["style"] == "--":
                horizontal_line.Line.DashStyle = 2  # 2 is the dash style for dashed lines

            message_top += message_spacing

    presentation.SaveAs(output_path)
    presentation.Close()
    PPTApp.Quit()
