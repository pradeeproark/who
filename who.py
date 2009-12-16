import win32com.client
import pywintypes
import pythoncom
import optparse
import sys
import time
import datetime
import win32api
import active_directory

selectedSettings = None
selectedMappings = None

try:
  import settings_mine
  selectedSettings = settings_mine
except:
  import settings 
  selectedSettings = settings

try:
  import mappings_mine
  selectedMappings = mappings_mine
except:
  import mappings
  selectedMappings = mappings

#Application GUID
_GUID = '{66c6759d-29dc-46a5-abab-b0f41df8328e}'
EventFlagIndexable   = 0x00000001
EventFlagHistorical  = 0x00000010

def fixMe(obj):
  if obj == None:
    return ''
  else:
    return unicode(obj)
    
def addContacts():
  try:
    event_factory = win32com.client.Dispatch('GoogleDesktopSearch.EventFactory')    
    import active_directory
    users = active_directory.AD_object (selectedSettings.ldapsearchpath)
    for user in users.search (objectCategory='Person'):
      try:
        event = event_factory.CreateEvent(_GUID, 'Google.Desktop.Contact')
        
        #Mandatory props
        event.AddProperty('format', 'text/html')
        event.AddProperty('content', '')
        event.AddProperty('last_modified_time', pywintypes.Time(time.time() + time.timezone))
        event.AddProperty('uri',u'%s%s' %(selectedSettings.uriprefix,fixMe(getattr(user,selectedMappings.uri))))
        atleastOneValidProp = False
        
        for prop in dir(selectedMappings):          
          if not prop.startswith('__') and prop <> 'uri' :
            print prop
            propval = getattr(selectedMappings,prop)          
            if propval <> '' :
              try:                               
                event.AddProperty(prop,fixMe(getattr(user,propval)))
                atleastOneValidProp = True
                print '.. got one prop'
              except:              
                print 'unable to add prop %s'   % prop     
                
        if atleastOneValidProp:
          event.Send(EventFlagIndexable)
          print '..done .. '
        else:
          print 'ignoring'
          
      except pywintypes.com_error, err:
        print err
    
  except pywintypes.com_error, e:
    print e    
  
def Main():
  
  parser = optparse.OptionParser(usage='%prog [options]')
  parser.add_option('-u', '--unregister', action='store_true', dest='unreg',
                    help='Run with this flag to unregister the plugin. '
                         'All other options are ignored when you use this flag.')
  (options, args) = parser.parse_args()  
  try:
    obj = win32com.client.Dispatch('GoogleDesktopSearch.Register')    
  except pythoncom.ole_error:
    print ('ERROR: You need to install Google Desktop Search to be able to '
           'use who.')
    sys.exit(2)

  if not options.unreg:
    try:
      # Register with GDS.  This is a one-time operation and will return an
      # error if already registered.  We cheat and just catch the error and
      # do nothing.
      
      # We try two different methods since different versions of GD have
      # different names for the registration method.  We ignore the specific
      # exception that we get if the method is not supported.
      try:
        obj.RegisterComponent(_GUID,['Title', 'who', 'Description', 'A contact and personal information indexer that '
                  'lets you index from active directory into your Google Desktop Search'
                  'index.', 'Icon', '%SystemRoot%\system32\SHELL32.dll,134'])
      except AttributeError, e:
        if (len(e.args) == 0 or
            e.args[0] != 'GoogleDesktopSearch.Register.RegisterComponent'):
          raise e
        else:
          obj.RegisterIndexingComponent(_GUID,['Title', 'who', 'Description', 'A contact and personal information indexer that '
                  'lets you index from active directory into your Google Desktop Search'
                  'index.', 'Icon', '%SystemRoot%\system32\SHELL32.dll,134'])          
    except pywintypes.com_error, e:
      if len(e.args) > 0 and e.args[0] == -2147352567:
        # This is the error we get if already registered.
        pass
      else:
        raise e
  else:
    # Try both approaches to unregister, too.
    try:
      obj.UnregisterIndexingComponent(_GUID)
    except:
      pass
    try:
      obj.UnregisterComponent(_GUID)
    except:
      pass
    sys.exit(0)
   
  addContacts()
if __name__ == '__main__':
  Main()
