# KMZ Processor
This is a collection of KMZ-related tools primarily built in Excel VBA. These tools are:

* **KMZ Separator:** Trims a provided KMZ to stay within the bounds of a Polygon drawn in a separate file
* **KMZ Address Counter**: Sums house placemarks within a KMZ, sorted by color
* **Address Separator**: Filters an "Address Validation" file, keeping only the addresses within a provided Polygon. Useful for isolating just the relevant addresses within a larger file
* **KMZ to CSV:** Extracts specific placemark info from a KMZ, reformats it to CSV. Intended to be used with ImportGeoCSV (included in lisp files in this repo) to place geolocated objects in AutoCAD. See below for a real-time demonstration of the macro in action
