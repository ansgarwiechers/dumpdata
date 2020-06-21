Dictionary-based data structures are handy in various situations. However, when the structures grow, they tend to become messy pretty fast.

To deal with this issue I wrote a function to inspect a given variable and return a string representation of its data. It was inspired by Perl's `Data::Dumper` module, although `DumpData()` is far less sophisticated (just in case you were wondering).
