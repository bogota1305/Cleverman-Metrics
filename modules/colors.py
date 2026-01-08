def lighten_color(color, factor=0.5):
    """Devuelve un color más claro basado en el color original."""
    color = color.lstrip('#')
    r, g, b = [int(color[i:i+2], 16) for i in (0, 2, 4)]
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"#{r:02X}{g:02X}{b:02X}"