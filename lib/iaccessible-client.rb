# iaccessible-client.rb
#
# A pure Ruby Microsoft Active Accessibility client.
#
# This module provides a way to use the MS Active Accessibility framework
# from Ruby.  Its goal is to wrap the IAccessible C interface in a set of
# Ruby objects and methods, without using any C extensions.
#
# Author: Bill Agee (billagee@gmail.com)
#
# More info:
#
# IAccessible
#   http://msdn.microsoft.com/en-us/library/dd318466%28v=vs.85%29.aspx
#
# Ye Olde Wikipedia article
#   http://en.wikipedia.org/wiki/Microsoft_Active_Accessibility

require 'rubygems'
require 'logger'
require 'win32/api'
require 'windows/com'
require 'windows/com/automation'
require 'windows/com/variant'
require 'windows/error'
require 'windows/msvcrt/buffer'
require 'windows/unicode'
require 'windows/window'
require 'windows/window/menu'

$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), 'lib'))
require 'iaccessible-client/accessibleobject'

include Windows::COM
include Windows::COM::Automation
include Windows::COM::Variant
include Windows::Error
include Windows::MSVCRT::Buffer
include Windows::Unicode
include Windows::Window
include Windows::Window::Menu


# C functions and constants from OleAcc.h
Wprintf = Windows::API.new('wprintf', 'PP', 'I', 'msvcrt')
AccessibleObjectFromWindow = Win32::API.new('AccessibleObjectFromWindow',
                                            'LLPP', 'L', 'oleacc')
AccessibleChildren = Win32::API.new('AccessibleChildren',
                                    'PLLPP', 'L', 'oleacc')
IID_IAccessible = [0x618736e0, 0x3c3d, 0x11cf, 0x81, 0x0c,
                   0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71].pack('LSSC8')
CHILDID_SELF = 0

module IAccessibleClient
  @@logger = Logger.new(STDOUT)
  @@logger.level = Logger::DEBUG
  #@@logger.level = Logger::WARN

  def self.logger
    @@logger
  end

  # Call CoInitialize() as soon as this module is loaded
  logger.debug("Calling CoInitialize to initialize COM")
  CoInitialize(nil)
  logger.debug("Using at_exit to register CoUninitialize call")
  at_exit { CoUninitialize() }


  # Returns an AccessibleObject instance that represents the desktop client
  # object.  This object's parent is the root desktop object, and its children
  # are the top-level windows on the system (such as open application windows).
  def self.desktop
    # Get the root desktop object
    desktop_root = get_desktop_root()
    # Get the desktop client object from the root, and return it
    desktop_client = desktop_root.find_child('Desktop')
    return desktop_client
  end

  # Returns an AccessibleObject for the desktop window (AKA the root element
  # in the UI tree).  Note that this is not a very useful object since the
  # top-level application windows on the system are children of the desktop
  # child object, which itself is a child of the desktop window.
  def self.get_desktop_root
    # Get the desktop's HWND
    desktop_hwnd = GetDesktopWindow()
    # Use the HWND to get an AccessibleObject for the desktop window
    desktop_root = AccessibleObject.new(:hwnd, desktop_hwnd)
    return desktop_root
  end

  # Takes an IAccessible object and returns the its name as a Ruby string.
  def self.get_iacc_object_name_as_ruby_str(iacc)
    hr, name_bstr = self.get_iacc_object_name_as_bstr(iacc)
    if (hr == 0)
      name = bstr_to_ruby_string(name_bstr)
    elsif (hr == 1)
      # This is the HRESULT value when the object has no name...pretty common
      # situation, so we simply return the string "(null)" just like get_accName
      # does.
      name = "(null)"
    else
      @@logger.error("Possible error!")
      @@logger.error("HRESULT from get_accName is: " + hr.to_s)
      @@logger.error("HRESULT message is: '#{get_last_error(hr)}'")
      # Anything else is probably an error, so raise an exception:
      raise "Error returned from get_accName while fetching object name!"
    end
    # Cleanup
    SysFreeString(name_bstr)
    return name
  end

  def self.bstr_to_ruby_string(bstr)
    # Get the BSTR length by reading the 4-byte length prefix located
    # immediately before the address of the BSTR.
    length_buf = 0.chr * 4
    memcpy(length_buf, bstr.unpack('L').first - 4, 4)
    # for some reason unpacking the raw bstr doesn't actually give us the
    # characters.  So, memcopy the BSTR's contents to a buffer and work with
    # the buffer instead:
    len = length_buf.unpack('L').first # Convert the LONG length to a Fixnum
    charbuf = 0.chr * len
    memcpy(charbuf, bstr.unpack('L').first, len)
    # Strip all nulls from the wide string. Without doing this, the Ruby
    # string looks like "f o o b a r "
    #
    # NOTE: Using a regexp here breaks frequently in Ruby 1.9.3 with:
    # `gsub': invalid byte sequence in US-ASCII (ArgumentError)
    #
    # So as of Ruby 1.9 one can't do:
    #return charbuf.gsub(/\000/, '')

    final_charbuf = String.new
    charbuf.each_char do |b|
      final_charbuf << b if (b != 0.chr)
    end
    return final_charbuf
  end
end

