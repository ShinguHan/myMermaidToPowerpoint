import win32com.client

def create_powerpoint_sequence_diagram(participants_info, messages_info, output_path):
    # Create PowerPoint application, presentation, and slide
    PPTApp = win32com.client.Dispatch("PowerPoint.Application")
    PPTApp.Visible = True
    presentation = PPTApp.Presentations.Add()
    slide = presentation.Slides.Add(1, 12)  # 12 is the layout index for blank slides

    # Add participants
    left = 100
    top = 100
    width = 100
    height = 50
    spacing = 150
    participant_shapes = []
    for participant_info in participants_info:
        shape = slide.Shapes.AddShape(1, left, top, width, height)
        shape.TextFrame.TextRange.Text = participant_info["text"]
        shape.Line.Visible = False
        shape.Fill.ForeColor.RGB = 0x000000
        participant_shapes.append(shape)

        # Draw vertical line for each participant
        vertical_line = slide.Shapes.AddLine(left + width/2, top + height, left + width/2, top + height + 400)
        vertical_line.Line.ForeColor.RGB = 0x000000

        left += spacing

    # Add messages
    message_top = top + height + 30
    message_spacing = 50
    for message_info in messages_info:
        from_participant = participant_shapes[message_info["from"]]
        to_participant = participant_shapes[message_info["to"]]

        # Add message text
        text_left = from_participant.Left + (abs(from_participant.Left - to_participant.Left) / 2) - 50
        text_shape = slide.Shapes.AddShape(1, text_left, message_top, 100, 20)
        text_shape.TextFrame.TextRange.Text = message_info["text"]
        text_shape.Line.Visible = False
        text_shape.Fill.ForeColor.RGB = 0x000000
        text_shape.TextFrame.TextRange.Font.Size = 10
        text_shape.TextFrame.AutoSize = 1

        # Add horizontal line with arrow
        horizontal_line = slide.Shapes.AddLine(from_participant.Left + width / 2, message_top + 10,
                                               to_participant.Left + width / 2, message_top + 10)
        horizontal_line.Line.ForeColor.RGB = 0x000000
        horizontal_line.Line.EndArrowheadStyle = 4  # 4 is the arrowhead style

        message_top += message_spacing

    presentation.SaveAs(output_path)
    presentation.Close()
    PPTApp.Quit()
