import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogBody,
  DialogActions,
  Button,
  Spinner,
} from "@fluentui/react-components";
import React from "react";

const SmallPopUp = (props) => {
  return (
    <Dialog {...props}>  {/* open={props.open} */}
      <DialogSurface>
        <DialogBody>
          {props.title && <DialogTitle>{props.title}</DialogTitle>}
          {(props.msg || props.children) && (
            <DialogContent>
              {props.spinner && (
                <div /* className={props.className} */>
                  <Spinner
                    size="huge"
                    labelPosition="below"
                    label={props.msg}
                  />
                </div>
              )}
              {props.children}
            </DialogContent>
          )}
          {
            /* props.open &&  */ !props.spinner && props.activeActions && (
              <DialogActions>
                <DialogTrigger disableButtonEnhancement>
                  <Button appearance="secondary">
                    {props.deleteTaskId ? "Cancel" : "Close"}
                  </Button>
                </DialogTrigger>
                {/* {props.deleteTaskId && (
                  <Button
                    appearance="primary"
                    // onClick={(e) => props.taskDelete(props.deleteTaskId)}
                  >
                    Delete
                  </Button>
                )} */}
              </DialogActions>
            )
          }
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default SmallPopUp;
