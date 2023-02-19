import pyproj

def itm_to_latlon(easting, northing):
    # Define the input and output coordinate reference systems
    itm_crs = pyproj.CRS('EPSG:2039')  # ITM (Israel Transverse Mercator)
    wgs84_crs = pyproj.CRS('EPSG:4326')  # WGS84 (latitude/longitude)

    # Create a PyProj transformer to convert coordinates from ITM to WGS84
    transformer = pyproj.Transformer.from_crs(itm_crs, wgs84_crs, always_xy=True)

    # Convert the input ITM coordinates to latitude and longitude
    longitude, latitude = transformer.transform(easting, northing)

    # Return the latitude and longitude as a tuple
    return latitude, longitude


lon, lat = itm_to_latlon(219000, 624000)
print(lon, lat )