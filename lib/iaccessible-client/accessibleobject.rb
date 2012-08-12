# accessibleobject.rb
#
# Defines the AccessibleObject class, a wrapper for IAccessible pointers.
#
# Author: Bill Agee (billagee@gmail.com)
#
# More info:
#
# IAccessible
#   http://msdn.microsoft.com/en-us/library/dd318466%28v=vs.85%29.aspx
#
# Some design aspects follow Perl's Win32::ActAcc module:
#   http://search.cpan.org/~pbwolf/Win32-ActAcc-1.1/

require 'rubygems'
require 'win32/api'
require 'windows/com'
require 'windows/com/automation'
require 'windows/com/variant'
require 'windows/error'
require 'windows/msvcrt/buffer'
require 'windows/unicode'
require 'windows/window'
require 'windows/window/menu'

include Windows::COM
include Windows::COM::Automation
include Windows::COM::Variant
include Windows::Error
include Windows::MSVCRT::Buffer
include Windows::Unicode
include Windows::Window
include Windows::Window::Menu

module IAccessibleClient
  class AccessibleObject
    @iacc = nil # The IAccessible* value for this object
    @iacc_ptr = nil # The memory address of the IAccessible
    attr_reader :iacc, :iacc_ptr, :name, :childcount
    
    def logger
      IAccessibleClient.logger
    end

    def initialize(how, what)
      case how
      when :iacc
        # Initialize object using an IAccessible pointer
        @iacc = what
      when :hwnd
        # Use the HWND to get the windows's IAccessible object
        hwnd = what
        @iacc = 0.chr * 4
        hr = AccessibleObjectFromWindow.call(hwnd, OBJID_WINDOW,
                                             IID_IAccessible, @iacc)
        raise "Failed to instantiate IAccessible object!" if (hr != S_OK)
      else
        raise "Invalid way ('#{how}') of finding accessible object!"
      end

      # Get the memory address of the object - need to use it to access the
      # object's virtual function table, and also use it as a this pointer:
      @iacc_ptr = @iacc.unpack('L').first

      logger.debug("Obtaining the data for the IAccessibleVtbl...")
      # Get the 28 function pointers that comprise the IAccessibleVtbl
      # C interface.  (See OleAcc.h for more details.)
      # We have to use memcpy to copy the IAccessible object into a buffer,
      # since we can't operate on the memory stored in @iacc_ptr.
      #
      # Storage for pointer to the vtbl
      lpvtbl = 0.chr * 4 
      # Storage for array of pointers to the IAccessible functions
      table = 0.chr * (4 * 28)
      # Copy the address from the IAccessible*
      memcpy(lpvtbl, @iacc_ptr, 4)
      # Copy the 112 bytes that make up the array of function pointers,
      # starting at the beginning of the IAccessible*
      memcpy(table, lpvtbl.unpack('L').first, 4 * 28)
      # Unpack the function pointers into a Ruby array. Each pointer is a LONG.
      vtbl_array = table.unpack('L*')

      logger.debug("Defining get_accName and other functions in the vtbl...")
      @get_accName = Win32::API::Function.new(vtbl_array[10], 'PLLLLP', 'L')
      @get_accChildCount = Win32::API::Function.new(vtbl_array[8], 'PP', 'L')
      @accDoDefaultAction = Win32::API::Function.new(vtbl_array[25],
                                                     'PLLLL', 'L')

      logger.debug("Getting child count of the current object...")
      childcount = 0.chr * 4
      childcount_ptr = [childcount].pack('P').unpack('L').first
      # TODO - Would this work instead, to simplify things a bit?
      #childcount_ptr = childcount.unpack('L').first
      hr = @get_accChildCount.call(@iacc_ptr, childcount_ptr)
      raise "get_accChildCount failed!" if (hr != 0)
      @childcount = childcount.unpack('L').first 
      logger.debug("Child count for this object is: #{@childcount}")

      logger.debug("Getting the name of the current object...")
      name_bstr = 0.chr * 4
      # The CHILDID_SELF VARIANT is packed as 4 unsigned shorts and two LONGs.
      # But when we pass it, it must be passed as four DWORDs:
      self_var = [VT_I4, 0, 0, 0, CHILDID_SELF, 0].pack('SSSSLL')
      # Could also try this for consistency with the function call below, but it
      # doesnt look like the tagVARIANT declaration, which could be confusing:
      #self_var = [VT_I4, 0, CHILDID_SELF, 0].pack('LLLL')

      #puts "Calling get_accName while init'ing object..."
      # NOTE! There's a frequent segfault here in Ruby 1.9.3p194
      # But it does not repro with Ruby 1.9.2p290.
      # https://github.com/oneclick/rubyinstaller/issues/116
      # See: http://bugs.ruby-lang.org/issues/6352
      # https://github.com/eventmachine/eventmachine/issues/333

      hr = @get_accName.call(@iacc_ptr,
                             self_var[0,4].unpack('L').first, # vt + resvd bytes
                             self_var[4,4].unpack('L').first, # 4 reserved bytes
                             self_var[8,4].unpack('L').first, # .bstrVal
                             self_var[12,4].unpack('L').first, # struct padding
                             name_bstr)
      if (hr == S_OK)
        # TODO - see about contributing bstr_to_ruby_string to windows-pr
        @name = IAccessibleClient.bstr_to_ruby_string(name_bstr)
      elsif (hr == 1)
        # An HRESULT of 1 means that the error wasn't fatal, but there's no name
        @name = "(null)"
      else
        #raise "HRESULT from get_accName indicates an error!"
        puts "While initializing, HRESULT from get_accName indicates an error!"
        @name = "(ERROR!)"
      end
      logger.debug("The name of the current object is '#{@name}'")
      return self
    end

    def find_child_recursively(name_wanted)
      sleep 0.5
      #puts "Beginning search for child '#{name_wanted}'..."
      iterate_accessible_children_recursively do |child|
        if (child.name == name_wanted)
          #puts "Found match for name_wanted: '#{child.name}'"
          return child
        #else
          # FIXME - release the child object
        end
      end
      raise "Loop complete, but failed to find child object '#{name_wanted}'!"
    end

    #def find_child(how, what) # FIXME - need to support :name and :id at least
    def find_child(name_wanted)
      #puts "Beginning search for child '#{name_wanted}'..."
      iterate_accessible_children do |child|
        if (child.name == name_wanted)
          #puts "Found match for name_wanted: '#{child.name}'"
          #puts child.inspect
          return child
        #else
          # FIXME - release the child object
        end
      end
      raise "Loop complete, but failed to find child object '#{name_wanted}'!"
    end

    def iterate_accessible_children_recursively
      padding = 0
      self.iterate_accessible_children do |child1|
        padding += 2
        print " " * padding
        puts "-----Recursing through children of child '#{child1.name}'-----"
        child1.iterate_accessible_children_recursively do |child2|
          puts child2.name
        end
        print " " * padding
        puts "-----Done Recursing children of child '#{child1.name}'-----"
      end
    end

    # Iterates over all accessible children of self, and yields to allow the
    # caller to execute a block using each child object.
    #
    # Note that "accessible children" means that only child objects with
    # an IDispatch pointer will be included in the iteration.  Child objects 
    # with no IDispatch (such as buttons with no children, and other objects
    # that are terminal nodes in the UI tree) must be accessed via their
    # parent element.
    def iterate_accessible_children
      # Allocate a pointer to LONG for the out param to give AccessibleChildren
      return_count = 0.chr * 4 
      # Allocate space for an array of VARIANTs.  AccessibleChildren wilhttp://search.cpan.org/~pbwolf/Win32-ActAcc-1.1/http://search.cpan.org/~pbwolf/Win32-ActAcc-1.1/l
      # place one VARIANT (16 bytes) into each array element.  Note that this
      # is a C array and not a Ruby array, so it will appear to be a single
      # binary string:
      child_variant_carray_buf = 0.chr * (16 * @childcount)
      #puts "CHILDCOUNT IS '#{@childcount}' - allocated a 16*#{@childcount} byte Ruby string:"
      #pp child_variant_carray_buf
      hr = AccessibleChildren.call(@iacc_ptr,
                                   0,
                                   @childcount,
                                   child_variant_carray_buf,
                                   return_count) 
      raise "AccessibleChildren failed!" if (hr != 0)

      #puts "carray buffer before split:"
      #pp child_variant_carray_buf

      return_count_unpacked = return_count.unpack('L').first
      #puts "Return count was '#{return_count_unpacked}'"

      # Split the packed buffer string into its individual VARIANTS, by using
      # map to get a Ruby array of 16-byte strings.  If this is successful
      # then each string will be a single VARIANT.

      # Old Ruby 1.8.7 way of building the array of child variants: 
#      child_variants = child_variant_carray_buf.scan(/.{16}/).map {|s| s}.flatten
      # NOTE: Interesting Ruby 1.9 fact - in 1.9, using scan with a regexp
      # on a string with packed binary data seems to always raise an exception.
      #
      # Instead split the packed string into an array of strings this way:
      child_variants = Array.new
      offset = 0
      return_count_unpacked.times do
        #puts "at offset '#{offset}'"
        child_variants << child_variant_carray_buf[offset, 16]
        offset += 16
      end

      #pp child_variants
      
      # Iterate over the children
      count = 1
      child_variants.each do |variant| 
        #puts "examining child variant #{count}"
        count += 1
        # We could unpack the entire variant into an array like this:
#        vtchild = variant.unpack('SSSSLL')
#        vt = vtchild[0]
        # Or, we can just access the members one at a time like this:
        vt = variant[0,2].unpack('S').first

        # Skip if the variant's .vt is not VT_DISPATCH. This avoids trying to QI
        # any variants that do not contain an IDispatch/IAccessible object.
        if (vt != VT_DISPATCH)
          #puts("No IDispatch/IAccessible in this VARIANT...skipping it.")
          next
        end
        # Get the IDispatch* value stored in the VARIANT's .pdispVal member.
        # This will be a Ruby Fixnum representing a memory address.
        pdispval_ptr = variant[8,4].unpack('L').first

        # We must get the QueryInterface function ptr of the IDispatch object,
        # so we can QI the IAccessible interface of the IDispatch.  Perhaps
        # there's some more graceful way to handle this; maybe make it a method?
        #
        # IDispatch contains 7 functions, so we need 7 blocks of 4 bytes for
        # each function's address:
        child_vtbl_ptr = 0.chr * 4
        child_vtbl = 0.chr * (4 * 7)
        memcpy(child_vtbl_ptr, pdispval_ptr, 4)
        memcpy(child_vtbl, child_vtbl_ptr.unpack('L').first, 4 * 7)
        child_vtbl = child_vtbl.unpack('L*')
        queryInterface = Win32::API::Function.new(child_vtbl[0], 'PPP', 'L')
        # Get the IAccessible of the IDispatch
        child_iacc = 0.chr * 4
        hr = queryInterface.call(pdispval_ptr, IID_IAccessible, child_iacc)
        raise "QueryInterface of child element failed!" if (hr != S_OK)

        child = AccessibleObject.new(:iacc, child_iacc)

        #puts "Child name is: '#{child.name}'"

        if block_given?
          yield child
        end

        # FIXME - whether a match is found or not, we need to release/free all
        # unneeded child objects and BSTRs created during the search...
      end
    end

    def print_child_names()
      puts "\nPrinting child names of element '#{@name}'...\n"
      iterate_accessible_children do |child|
        puts "Child name is: '#{child.name}'"
      end
    end

    # Prints the name of the IAccessible object using wprintf.
    def print_name_wprintf()
      hr, name_bstr = get_iacc_object_name_as_bstr(@iacc)
      if (hr == S_OK)
        # Create format string
        format_str = multi_to_wide("The BSTR printed with wprintf is: '%s'\n")
        # Pack the format string as a C-style string and save its address
        format_str_ptr = [format_str].pack('p').unpack('L').first
        # Pass the format string pointer and bstr pointer to wprintf
        Wprintf.call(format_str_ptr, name_bstr.unpack('L').first)
      else
        puts "(null)"
      end
      SysFreeString(name_bstr)
    end

    def do_default_action(child_id=CHILDID_SELF)
      self_var = [VT_I4, 0, 0, 0, child_id, 0].pack('SSSSLL')
      hr = @accDoDefaultAction.call(@iacc_ptr,
                                    self_var[0,4].unpack('L').first,
                                    self_var[4,4].unpack('L').first,
                                    self_var[8,4].unpack('L').first,
                                    self_var[12,4].unpack('L').first)
      if (hr != S_OK)
        puts "ERROR: HRESULT from accDoDefaultAction is: " + hr.to_s
        puts "Th HRESULT message for that error is: '#{get_last_error(hr)}'"
        raise "accDoDefaultAction did not return success!"
      end
    end
  end
end

